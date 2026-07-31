import os
import ssl
import warnings
import winsound
import numpy as np
import easyocr
import keyboard
import pyperclip
from PIL import Image, ImageGrab, ImageEnhance

# 1. Suppress PyTorch warnings and bypass SSL
warnings.filterwarnings("ignore", category=UserWarning)
ssl._create_default_https_context = ssl._create_unverified_context

# 2. Initialize EasyOCR reader once globally
print("Loading OCR models... Please wait.")
reader = easyocr.Reader(['en'], gpu=False)
print("OCR Tool is ready! Press 'Ctrl + Shift + O' to extract text.")
print("Press 'Ctrl + Shift + X' to completely exit the script.")

def enhance_image(pil_img):
    """Sharpens and boosts contrast for low-res or stylized fonts."""
    img_gray = pil_img.convert('L')
    contrast = ImageEnhance.Contrast(img_gray)
    img_contrast = contrast.enhance(2.0)
    sharpness = ImageEnhance.Sharpness(img_contrast)
    img_sharp = sharpness.enhance(2.0)
    return img_sharp

def process_clipboard_ocr():
    print("\nHOTKEY DETECTED: Processing...")
    
    # 3. Grab data from clipboard
    clipboard_data = ImageGrab.grabclipboard()

    if clipboard_data is not None:
        try:
            # 4. FIX: Handle case where clipboard returns a list of file paths
            if isinstance(clipboard_data, list):
                if len(clipboard_data) > 0 and os.path.exists(clipboard_data[0]):
                    # Open the first image file path in the list
                    img = Image.open(clipboard_data[0])
                else:
                    print("Clipboard list does not contain a valid file path.")
                    winsound.Beep(400, 300)
                    return
            else:
                # Clipboard contains raw image data directly
                img = clipboard_data

            # 5. Sharpen and process the image
            processed_img = enhance_image(img)
            
            # 6. Convert PIL Image to NumPy array
            img_np = np.array(processed_img)
            
            # 7. Perform OCR
            results = reader.readtext(img_np, detail=0)
            extracted_text = " ".join(results)
            
            if extracted_text.strip():
                # 8. Automatically copy text back to clipboard
                pyperclip.copy(extracted_text)
                print(f"SUCCESS! Copied to clipboard: {extracted_text}")
                
                # Double beep for success
                winsound.Beep(1000, 100)
                winsound.Beep(1200, 100)
            else:
                print("OCR ran but found no readable text.")
                winsound.Beep(400, 300)
                
        except Exception as e:
            print(f"An error occurred inside processing: {e}")
            winsound.Beep(400, 300)
    else:
        print("Error: Clipboard is empty or does not contain an image.")
        winsound.Beep(400, 300)

if __name__ == "__main__":
    # Register the main OCR shortcut
    keyboard.add_hotkey('ctrl+shift+q', process_clipboard_ocr)
    
    # Clean exit shortcut so you don't have to break it manually in VS Code
    keyboard.add_hotkey('ctrl+shift+x', os._exit, args=(0,))
    
    # Keep the script running
    keyboard.wait()
