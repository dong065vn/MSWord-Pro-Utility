import os
import sys
from PIL import Image, ImageDraw

def create_rounded_mask(size, radius):
    mask = Image.new("L", size, 0)
    draw = ImageDraw.Draw(mask)
    draw.rounded_rectangle([(0, 0), size], radius, fill=255)
    return mask

def process_image(img_path, output_dir):
    # Load original image
    img = Image.open(img_path).convert("RGBA")
    
    # Ensure square image by cropping center if needed
    width, height = img.size
    min_dim = min(width, height)
    left = (width - min_dim) / 2
    top = (height - min_dim) / 2
    right = (width + min_dim) / 2
    bottom = (height + min_dim) / 2
    img = img.crop((left, top, right, bottom))
    
    sizes = [16, 24, 32, 48, 64, 128, 256, 512, 1024]
    
    print("Generating square and rounded icons...")
    
    icon_images_rounded = []
    icon_images_square = []
    
    for size in sizes:
        # Square version
        resized_square = img.resize((size, size), Image.Resampling.LANCZOS)
        resized_square.save(os.path.join(output_dir, f"icon_square_{size}x{size}.png"), format="PNG")
        
        # Rounded version
        radius = int(size * 0.2)  # 20% rounding
        mask = create_rounded_mask((size, size), radius)
        
        # Create transparent background for rounded corners
        bg = Image.new('RGBA', (size, size), (0,0,0,0))
        bg.paste(resized_square, (0, 0), mask)
        
        bg.save(os.path.join(output_dir, f"icon_rounded_{size}x{size}.png"), format="PNG")
        
        if size <= 256:
            icon_images_rounded.append(bg)
            icon_images_square.append(resized_square)
            
    # Save .ico files
    print("Saving .ico files...")
    ico_sizes = [(16,16), (24,24), (32,32), (48,48), (64,64), (128,128), (256,256)]
    img.save(os.path.join(output_dir, "app_icon_square.ico"), format="ICO", sizes=ico_sizes)
    
    # For rounded ico, we use the 1024x1024 rounded image as base
    if icon_images_rounded:
        icon_images_rounded[-1].save(os.path.join(output_dir, "app_icon_rounded.ico"), format="ICO", sizes=ico_sizes)
        
    print(f"Icons generated successfully in {output_dir}")

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python generate_icons.py <input_image_path>")
        sys.exit(1)
        
    input_path = sys.argv[1]
    output_dir = os.path.dirname(os.path.abspath(__file__))
    
    if not os.path.exists(input_path):
        print(f"Error: Could not find image at {input_path}")
        sys.exit(1)
        
    process_image(input_path, output_dir)
