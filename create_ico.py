from PIL import Image
import os

def create_ico():
    # Check if icons directory exists
    if not os.path.exists('icons'):
        print("Icons directory not found!")
        return False

    # Find the PNG icon
    png_path = os.path.join('icons', 'icon.png')
    if not os.path.exists(png_path):
        print("icon.png not found in icons directory!")
        return False

    # Open the PNG image
    img = Image.open(png_path)

    # Create ICO file
    ico_path = os.path.join('icons', 'icon.ico')
    
    # Convert and save as ICO
    # ICO format typically includes multiple sizes
    # Common sizes are 16x16, 32x32, 48x48, and 256x256
    sizes = [(16,16), (32,32), (48,48), (256,256)]
    
    img.save(ico_path, format='ICO', sizes=sizes)
    print(f"Created {ico_path}")
    return True

if __name__ == "__main__":
    create_ico() 