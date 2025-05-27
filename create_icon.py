import cairosvg
from PIL import Image
import io

def create_icon():
    # Convert SVG to PNG in memory with high resolution
    png_data = cairosvg.svg2png(url="icon.svg", output_width=256, output_height=256)
    
    # Create PIL Image from PNG data
    img = Image.open(io.BytesIO(png_data))
    
    # Save as PNG
    img.save("icon.png", format='PNG')
    
    # Save as ICO with multiple sizes
    img.save("icon.ico", format='ICO', sizes=[(16, 16), (32, 32), (48, 48), (256, 256)])

if __name__ == "__main__":
    create_icon() 