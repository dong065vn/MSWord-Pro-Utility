param(
    [string]$RepoRoot = (Split-Path -Parent $PSScriptRoot),
    [string]$Version
)

$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Drawing
Add-Type @"
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;

public static class RepoExeIconPatcher
{
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    public struct SHFILEINFO
    {
        public IntPtr hIcon;
        public IntPtr iIcon;
        public uint dwAttributes;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)] public string szDisplayName;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 80)] public string szTypeName;
    }

    private const int RT_ICON = 3;
    private const int RT_GROUP_ICON = 14;
    private const uint LOAD_LIBRARY_AS_DATAFILE = 0x00000002;
    public const uint SHGFI_ICON = 0x000000100;
    public const uint SHGFI_SMALLICON = 0x000000001;
    public const uint SHGFI_LARGEICON = 0x000000000;

    [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    private static extern IntPtr BeginUpdateResourceW(string pFileName, bool bDeleteExistingResources);
    [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    private static extern bool UpdateResourceW(IntPtr hUpdate, IntPtr lpType, IntPtr lpName, ushort wLanguage, byte[] lpData, uint cbData);
    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool EndUpdateResourceW(IntPtr hUpdate, bool fDiscard);
    [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    private static extern IntPtr LoadLibraryExW(string lpFileName, IntPtr hReservedNull, uint dwFlags);
    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool FreeLibrary(IntPtr hModule);
    [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
    public static extern IntPtr SHGetFileInfo(string pszPath, uint dwFileAttributes, out SHFILEINFO psfi, uint cbFileInfo, uint uFlags);
    [DllImport("user32.dll", SetLastError = true)]
    public static extern bool DestroyIcon(IntPtr hIcon);
    private delegate bool EnumResNameProc(IntPtr hModule, IntPtr lType, IntPtr lName, IntPtr lParam);
    [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    private static extern bool EnumResourceNamesW(IntPtr hModule, IntPtr lType, EnumResNameProc lpEnumFunc, IntPtr lParam);

    private static IntPtr MakeIntResource(int id) => (IntPtr)id;
    private static bool IsIntResource(IntPtr ptr) => (((ulong)ptr.ToInt64()) >> 16) == 0;
    private static int ResourceId(IntPtr ptr) => ptr.ToInt32() & 0xFFFF;

    private sealed class IconImage
    {
        public int Width;
        public int Height;
        public byte ColorCount;
        public ushort Planes;
        public ushort BitCount;
        public byte[] Bytes = Array.Empty<byte>();
    }

    public static int[] GetGroupIconIds(string exePath)
    {
        var module = LoadLibraryExW(exePath, IntPtr.Zero, LOAD_LIBRARY_AS_DATAFILE);
        if (module == IntPtr.Zero) throw new InvalidOperationException("LoadLibraryExW failed: " + Marshal.GetLastWin32Error());
        try
        {
            var ids = new List<int>();
            EnumResNameProc callback = (hModule, lType, lName, lParam) => { if (IsIntResource(lName)) ids.Add(ResourceId(lName)); return true; };
            if (!EnumResourceNamesW(module, MakeIntResource(RT_GROUP_ICON), callback, IntPtr.Zero))
            {
                var error = Marshal.GetLastWin32Error();
                if (error != 1813) throw new InvalidOperationException("EnumResourceNamesW failed: " + error);
            }
            return ids.Count == 0 ? new[] { 1 } : ids.Distinct().ToArray();
        }
        finally { FreeLibrary(module); }
    }

    public static void PatchIcon(string exePath, string iconPath)
    {
        var entries = ParseIco(iconPath);
        var groupIds = GetGroupIconIds(exePath);
        var handle = BeginUpdateResourceW(exePath, false);
        if (handle == IntPtr.Zero) throw new InvalidOperationException("BeginUpdateResourceW failed: " + Marshal.GetLastWin32Error());
        try
        {
            for (int i = 0; i < entries.Count; i++)
            {
                var id = i + 1;
                if (!UpdateResourceW(handle, MakeIntResource(RT_ICON), MakeIntResource(id), 0, entries[i].Bytes, (uint)entries[i].Bytes.Length))
                    throw new InvalidOperationException("UpdateResourceW RT_ICON failed: " + Marshal.GetLastWin32Error());
            }
            var groupData = BuildGroupIcon(entries);
            foreach (var groupId in groupIds)
            {
                if (!UpdateResourceW(handle, MakeIntResource(RT_GROUP_ICON), MakeIntResource(groupId), 0, groupData, (uint)groupData.Length))
                    throw new InvalidOperationException("UpdateResourceW RT_GROUP_ICON failed: " + Marshal.GetLastWin32Error());
            }
        }
        catch
        {
            EndUpdateResourceW(handle, true);
            throw;
        }
        if (!EndUpdateResourceW(handle, false)) throw new InvalidOperationException("EndUpdateResourceW failed: " + Marshal.GetLastWin32Error());
    }

    private static List<IconImage> ParseIco(string path)
    {
        using var fs = File.OpenRead(path);
        using var br = new BinaryReader(fs);
        var reserved = br.ReadUInt16();
        var type = br.ReadUInt16();
        var count = br.ReadUInt16();
        if (reserved != 0 || type != 1 || count == 0) throw new InvalidDataException("Invalid ICO header");
        var headers = new List<(byte Width, byte Height, byte ColorCount, ushort Planes, ushort BitCount, uint Size, uint Offset)>();
        for (int i = 0; i < count; i++)
        {
            var width = br.ReadByte();
            var height = br.ReadByte();
            var colorCount = br.ReadByte();
            br.ReadByte();
            var planes = br.ReadUInt16();
            var bitCount = br.ReadUInt16();
            var size = br.ReadUInt32();
            var offset = br.ReadUInt32();
            headers.Add((width, height, colorCount, planes, bitCount, size, offset));
        }
        var result = new List<IconImage>();
        foreach (var h in headers)
        {
            fs.Position = h.Offset;
            var bytes = br.ReadBytes((int)h.Size);
            result.Add(new IconImage { Width = h.Width == 0 ? 256 : h.Width, Height = h.Height == 0 ? 256 : h.Height, ColorCount = h.ColorCount, Planes = h.Planes, BitCount = h.BitCount, Bytes = bytes });
        }
        return result;
    }

    private static byte[] BuildGroupIcon(List<IconImage> entries)
    {
        using var ms = new MemoryStream();
        using var bw = new BinaryWriter(ms);
        bw.Write((ushort)0);
        bw.Write((ushort)1);
        bw.Write((ushort)entries.Count);
        for (int i = 0; i < entries.Count; i++)
        {
            var entry = entries[i];
            bw.Write((byte)(entry.Width >= 256 ? 0 : entry.Width));
            bw.Write((byte)(entry.Height >= 256 ? 0 : entry.Height));
            bw.Write(entry.ColorCount);
            bw.Write((byte)0);
            bw.Write(entry.Planes);
            bw.Write(entry.BitCount);
            bw.Write((uint)entry.Bytes.Length);
            bw.Write((ushort)(i + 1));
        }
        return ms.ToArray();
    }
}
"@

function Get-VersionFromConfig([string]$ConfigPath) {
    $content = Get-Content $ConfigPath -Raw
    $match = [regex]::Match($content, 'APP_VERSION\s*=\s*"([^"]+)"')
    if (-not $match.Success) { throw "Could not determine APP_VERSION from $ConfigPath" }
    $match.Groups[1].Value
}

function Get-LatestVersion([string]$Repo, [string]$ConfigPath) {
    $versions = New-Object System.Collections.Generic.List[version]
    try { $versions.Add([version](Get-VersionFromConfig $ConfigPath)) } catch {}
    Get-ChildItem (Join-Path $Repo 'Releases') -Filter 'RELEASE_NOTES_v*.md' -File -ErrorAction SilentlyContinue | ForEach-Object {
        if ($_.BaseName -match '^RELEASE_NOTES_v(\d+\.\d+\.\d+)$') {
            try { $versions.Add([version]$matches[1]) } catch {}
        }
    }
    if ($versions.Count -eq 0) { throw 'Could not determine existing version.' }
    ($versions | Sort-Object -Descending | Select-Object -First 1).ToString()
}

function Get-NextPatchVersion([string]$VersionText) {
    $v = [version]$VersionText
    '{0}.{1}.{2}' -f $v.Major, $v.Minor, ($v.Build + 1)
}

function Set-VersionInFiles([string]$ConfigPath, [string]$MainPath, [string]$NewVersion) {
    $configContent = Get-Content $ConfigPath -Raw
    $configUpdated = [regex]::Replace($configContent, '(?m)(Global Const \$VERSION = ")([^"]+)(")', ('${1}' + $NewVersion + '${3}'), 1)
    $configUpdated = [regex]::Replace($configUpdated, '(?m)(Global Const \$APP_VERSION = ")([^"]+)(")', ('${1}' + $NewVersion + '${3}'), 1)
    if ($configUpdated -eq $configContent) { $configUpdated = $configContent }
    Set-Content $ConfigPath -Value $configUpdated -Encoding UTF8

    $mainContent = Get-Content $MainPath -Raw
    $mainUpdated = [regex]::Replace($mainContent, '(?m)(; PDF to Word Fixer Pro v)(\d+\.\d+\.\d+)( - MODULAR ARCHITECTURE)', ('${1}' + $NewVersion + '${3}'), 1)
    if ($mainUpdated -eq $mainContent) { $mainUpdated = $mainContent }
    Set-Content $MainPath -Value $mainUpdated -Encoding UTF8
}

function Get-IconPngSet([string]$BaseDir, [string]$Prefix) {
    $expectedSizes = 16, 24, 32, 48, 64, 128, 256, 512, 1024
    $items = foreach ($size in $expectedSizes) {
        $path = Join-Path $BaseDir ('{0}_{1}x{1}.png' -f $Prefix, $size)
        if (-not (Test-Path $path)) { throw "Missing icon asset: $path" }
        [PSCustomObject]@{ Path = $path; Size = $size }
    }
    ,$items
}

function New-IcoFromPngSet($IconSet, [string]$IcoPath) {
    $entries = foreach ($item in $IconSet | Where-Object Size -le 256) {
        [PSCustomObject]@{ Size = $item.Size; Bytes = [System.IO.File]::ReadAllBytes($item.Path) }
    }
    $fs = [System.IO.File]::Open($IcoPath, [System.IO.FileMode]::Create)
    $bw = New-Object System.IO.BinaryWriter($fs)
    try {
        $bw.Write([UInt16]0)
        $bw.Write([UInt16]1)
        $bw.Write([UInt16]$entries.Count)
        $offset = 6 + (16 * $entries.Count)
        foreach ($entry in $entries) {
            $dim = if ($entry.Size -ge 256) { 0 } else { $entry.Size }
            $bw.Write([byte]$dim)
            $bw.Write([byte]$dim)
            $bw.Write([byte]0)
            $bw.Write([byte]0)
            $bw.Write([UInt16]1)
            $bw.Write([UInt16]32)
            $bw.Write([UInt32]$entry.Bytes.Length)
            $bw.Write([UInt32]$offset)
            $offset += $entry.Bytes.Length
        }
        foreach ($entry in $entries) { $bw.Write($entry.Bytes) }
    }
    finally {
        $bw.Close()
        $fs.Close()
    }
}

function Render-BytesHash([byte[]]$Bytes) {
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try { return ([BitConverter]::ToString($sha.ComputeHash($Bytes))).Replace('-', '') }
    finally { $sha.Dispose() }
}

function Render-IconHash([System.Drawing.Icon]$Icon, [int]$Size) {
    $bmp = New-Object System.Drawing.Bitmap($Size, $Size)
    $graphics = [System.Drawing.Graphics]::FromImage($bmp)
    $graphics.Clear([System.Drawing.Color]::Transparent)
    $graphics.DrawIcon($Icon, 0, 0)
    $ms = New-Object System.IO.MemoryStream
    $bmp.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png)
    $graphics.Dispose()
    $bmp.Dispose()
    $bytes = $ms.ToArray()
    $ms.Dispose()
    Render-BytesHash -Bytes $bytes
}

function Render-ImageFileHash([string]$Path, [int]$Size) {
    $img = [System.Drawing.Image]::FromFile($Path)
    $bmp = New-Object System.Drawing.Bitmap($Size, $Size)
    $graphics = [System.Drawing.Graphics]::FromImage($bmp)
    $graphics.Clear([System.Drawing.Color]::Transparent)
    $graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $graphics.DrawImage($img, 0, 0, $Size, $Size)
    $ms = New-Object System.IO.MemoryStream
    $bmp.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png)
    $graphics.Dispose()
    $bmp.Dispose()
    $img.Dispose()
    $bytes = $ms.ToArray()
    $ms.Dispose()
    Render-BytesHash -Bytes $bytes
}

function Invoke-PatchIconWithRetry([string]$ExePath, [string]$IconPath, [int]$MaxAttempts = 8) {
    $lastError = $null
    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        try {
            [RepoExeIconPatcher]::PatchIcon($ExePath, $IconPath)
            return
        }
        catch {
            $lastError = $_
            Start-Sleep -Milliseconds (400 * $attempt)
        }
    }
    throw $lastError
}

function Sync-IconAssets([string]$Repo) {
    $iconDir = Join-Path $Repo 'app_icons'
    $resourcesDir = Join-Path $Repo 'Resources'
    $roundedSet = Get-IconPngSet -BaseDir $iconDir -Prefix 'icon_rounded'
    $squareSet = Get-IconPngSet -BaseDir $iconDir -Prefix 'icon_square'
    $roundedIco = Join-Path $iconDir 'app_icon_rounded.ico'
    $squareIco = Join-Path $iconDir 'app_icon_square.ico'
    New-IcoFromPngSet -IconSet $roundedSet -IcoPath $roundedIco
    New-IcoFromPngSet -IconSet $squareSet -IcoPath $squareIco
    Copy-Item $roundedIco (Join-Path $resourcesDir 'icon.ico') -Force
    Copy-Item ((($roundedSet | Where-Object Size -eq 32).Path)) (Join-Path $resourcesDir 'icon.png') -Force
    [PSCustomObject]@{
        RoundedSet = $roundedSet
        RoundedIco = $roundedIco
        SquareIco = $squareIco
        ResourceIco = (Join-Path $resourcesDir 'icon.ico')
        ResourcePng = (Join-Path $resourcesDir 'icon.png')
    }
}

function Copy-CurrentSourceTree([string]$Repo, [string]$DestinationRoot) {
    Get-ChildItem $Repo -Recurse -File -Force | ForEach-Object {
        $relative = $_.FullName.Substring($Repo.Length).TrimStart('\')
        if ($relative -match '^(?:\.git|\.vs|bin|obj)(\\|$)') { return }
        if ($relative -match '^Releases\\v[^\\]+\\binary(\\|$)') { return }
        if ($relative -match '^Releases\\PDF_to_Word_Fixer_Pro_v.+\.zip$') { return }
        if ($relative -match '^Main_compiled(?:_v\d+\.\d+\.\d+)?\.exe$') { return }
        if ($relative -match '^Main_build_\d+.*\.exe$') { return }
        if ($relative -match '^aut2exe_build\.log$') { return }
        $dest = Join-Path $DestinationRoot $relative
        $destDir = Split-Path $dest -Parent
        if (-not (Test-Path $destDir)) { New-Item -ItemType Directory -Force -Path $destDir | Out-Null }
        Copy-Item $_.FullName $dest -Recurse -Force
    }
}

function Ensure-ReleaseNotes([string]$NotesPath, [string]$VersionValue) {
    if (Test-Path $NotesPath) { return }
    @"
# PDF to Word Fixer Pro v$VersionValue

Release date: $(Get-Date -Format 'yyyy-MM-dd')

Highlights:
- Rebuilt and synchronized all app logo assets from the full rounded and square PNG icon sets.
- Build pipeline now regenerates `.ico` files and syncs `Resources` before compile to avoid logo build drift.

Validation:
- Au3Check clean for Main.au3
- Embedded logo verification passed for Explorer icon sizes
"@ | Set-Content $NotesPath -Encoding UTF8
}

function New-Readme([string]$ReadmePath, [string]$VersionValue) {
    @"
# PDF to Word Fixer Pro v$VersionValue

Binary package for release v$VersionValue.

Included:
- Main_compiled.exe
- RELEASE_NOTES_v$VersionValue.md
- app_icons\*
"@ | Set-Content $ReadmePath -Encoding UTF8
}

$repo = (Resolve-Path $RepoRoot).Path
$configPath = Join-Path $repo 'Config.au3'
$mainPath = Join-Path $repo 'Main.au3'
$rootExePath = Join-Path $repo 'Main_compiled.exe'
$aut2exe = 'C:\Program Files (x86)\AutoIt3\Aut2Exe\Aut2exe.exe'
$au3check = 'C:\Program Files (x86)\AutoIt3\Au3Check.exe'
$latestVersion = Get-LatestVersion -Repo $repo -ConfigPath $configPath
$targetVersion = if ($Version) { $Version } else { Get-NextPatchVersion $latestVersion }
$versionedExePath = Join-Path $repo ("Main_compiled_v{0}.exe" -f $targetVersion)
$releaseDir = Join-Path $repo ("Releases\v{0}" -f $targetVersion)
$releaseBinaryDir = Join-Path $releaseDir 'binary'
$binaryZip = Join-Path $repo ("Releases\PDF_to_Word_Fixer_Pro_v{0}_binary.zip" -f $targetVersion)
$sourceZip = Join-Path $repo ("Releases\PDF_to_Word_Fixer_Pro_v{0}_source.zip" -f $targetVersion)
$notesPath = Join-Path $repo ("Releases\RELEASE_NOTES_v{0}.md" -f $targetVersion)
$binaryStage = Join-Path $env:TEMP ("pdf_to_word_fixer_binary_{0}" -f $targetVersion)
$sourceStage = Join-Path $env:TEMP ("pdf_to_word_fixer_source_{0}" -f $targetVersion)

if (-not (Test-Path $aut2exe)) { throw "Missing Aut2Exe: $aut2exe" }
if (-not (Test-Path $au3check)) { throw "Missing Au3Check: $au3check" }

Write-Host '[1/7] Sync logo assets'
$iconState = Sync-IconAssets -Repo $repo

Write-Host '[2/7] Update version'
Set-VersionInFiles -ConfigPath $configPath -MainPath $mainPath -NewVersion $targetVersion

Write-Host '[3/7] Au3Check'
& $au3check $mainPath
if ($LASTEXITCODE -ne 0) { throw "Au3Check failed with exit code $LASTEXITCODE" }

Write-Host '[4/7] Compile'
Remove-Item $rootExePath -Force -ErrorAction SilentlyContinue
Remove-Item $versionedExePath -Force -ErrorAction SilentlyContinue
& $aut2exe /in $mainPath /out $rootExePath
Start-Sleep -Milliseconds 1000
if (-not (Test-Path $rootExePath)) { throw "Aut2Exe did not produce $rootExePath" }

Write-Host '[5/7] Patch and verify embedded icon'
Invoke-PatchIconWithRetry -ExePath $rootExePath -IconPath $iconState.RoundedIco
$icon32Path = ($iconState.RoundedSet | Where-Object Size -eq 32 | Select-Object -ExpandProperty Path)
$smallInfo = New-Object RepoExeIconPatcher+SHFILEINFO
[RepoExeIconPatcher]::SHGetFileInfo($rootExePath, 0, [ref]$smallInfo, [System.Runtime.InteropServices.Marshal]::SizeOf([type]([RepoExeIconPatcher+SHFILEINFO])), [RepoExeIconPatcher]::SHGFI_ICON -bor [RepoExeIconPatcher]::SHGFI_SMALLICON) | Out-Null
$smallIcon = [System.Drawing.Icon]::FromHandle($smallInfo.hIcon)
$exe16 = Render-IconHash -Icon $smallIcon -Size 16
[RepoExeIconPatcher]::DestroyIcon($smallInfo.hIcon) | Out-Null
$largeInfo = New-Object RepoExeIconPatcher+SHFILEINFO
[RepoExeIconPatcher]::SHGetFileInfo($rootExePath, 0, [ref]$largeInfo, [System.Runtime.InteropServices.Marshal]::SizeOf([type]([RepoExeIconPatcher+SHFILEINFO])), [RepoExeIconPatcher]::SHGFI_ICON -bor [RepoExeIconPatcher]::SHGFI_LARGEICON) | Out-Null
$largeIcon = [System.Drawing.Icon]::FromHandle($largeInfo.hIcon)
$exe32 = Render-IconHash -Icon $largeIcon -Size 32
[RepoExeIconPatcher]::DestroyIcon($largeInfo.hIcon) | Out-Null
$png16 = Render-ImageFileHash -Path ((($iconState.RoundedSet | Where-Object Size -eq 16).Path)) -Size 16
$png32 = Render-ImageFileHash -Path $icon32Path -Size 32
if ($png16 -ne $exe16) { throw "Embedded icon 16px mismatch. PNG=$png16 EXE=$exe16" }
if ($png32 -ne $exe32) { throw "Embedded icon 32px mismatch. PNG=$png32 EXE=$exe32" }

Write-Host '[6/7] Stage release files'
Ensure-ReleaseNotes -NotesPath $notesPath -VersionValue $targetVersion
Remove-Item $binaryStage -Recurse -Force -ErrorAction SilentlyContinue
New-Item -ItemType Directory -Force -Path $binaryStage | Out-Null
New-Readme -ReadmePath (Join-Path $binaryStage 'README.md') -VersionValue $targetVersion
Copy-Item $rootExePath $versionedExePath -Force
Copy-Item $rootExePath (Join-Path $binaryStage 'Main_compiled.exe') -Force
Copy-Item $notesPath (Join-Path $binaryStage (Split-Path $notesPath -Leaf)) -Force
Copy-Item (Join-Path $repo 'app_icons') (Join-Path $binaryStage 'app_icons') -Recurse -Force
Copy-Item (Join-Path $repo 'Resources') (Join-Path $binaryStage 'Resources') -Recurse -Force
New-Item -ItemType Directory -Force -Path $releaseBinaryDir | Out-Null
Copy-Item (Join-Path $binaryStage '*') $releaseBinaryDir -Recurse -Force

Write-Host '[7/7] Package zip assets'
Remove-Item $binaryZip -Force -ErrorAction SilentlyContinue
Remove-Item $sourceZip -Force -ErrorAction SilentlyContinue
Compress-Archive -Path (Join-Path $binaryStage '*') -DestinationPath $binaryZip -CompressionLevel Optimal
Remove-Item $sourceStage -Recurse -Force -ErrorAction SilentlyContinue
New-Item -ItemType Directory -Force -Path $sourceStage | Out-Null
Copy-CurrentSourceTree -Repo $repo -DestinationRoot $sourceStage
Compress-Archive -Path (Join-Path $sourceStage '*') -DestinationPath $sourceZip -CompressionLevel Optimal

Write-Host ''
Write-Host "Build completed successfully for v$targetVersion"
Write-Host "Rounded ICO: $($iconState.RoundedIco)"
Write-Host "Square ICO: $($iconState.SquareIco)"
Write-Host "Resource ICO: $($iconState.ResourceIco)"
Write-Host "Resource PNG: $($iconState.ResourcePng)"
Write-Host "Main exe: $rootExePath"
Write-Host "Versioned exe: $versionedExePath"
Write-Host "Binary zip: $binaryZip"
Write-Host "Source zip: $sourceZip"

