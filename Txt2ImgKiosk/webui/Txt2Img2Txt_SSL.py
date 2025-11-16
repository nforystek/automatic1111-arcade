from PIL import Image
from datetime import datetime

import certifi
import urllib3
import requests
import base64
import sys
import uu
import io

# Define the API URL and payload
url = "https://127.0.0.1:7860/sdapi/v1/txt2img"
payload = {
    "seed": sys.argv[1],
    "steps": sys.argv[2],
    "height": sys.argv[3],
    "width": sys.argv[4],
    "prompt": sys.argv[5],
    "negative_prompt": sys.argv[6]
}

http = urllib3.PoolManager(
    cert_reqs="CERT_REQUIRED"
)

# Send the request
response = http.request("POST", url, json=payload, timeout=120)
result = response.json()

# Encoding the data
def uuencode_data(input_data):
    # Create a BytesIO object to hold the encoded data
    encoded_buffer = io.BytesIO()
    
    # Use uu.encode to encode the data into the buffer
    with io.BytesIO(input_data) as input_buffer:
        uu.encode(input_buffer, encoded_buffer, name=str(datetime.utcnow()).replace("-","").replace(":","").replace(" ","").replace('.',"") + ".bmp")
   
    # Retrieve the encoded data as a string
    encoded_buffer.seek(0)
    return encoded_buffer.read().decode()

def convert_png_to_bmp(png_data):
    """
    Converts PNG image data to BMP format.

    Args:
        png_data (bytes): The PNG image data as a byte stream.

    Returns:
        bytes: The BMP image data as a byte stream.
    """
    try:
        # Open the PNG image from the byte stream
        png_image = Image.open(io.BytesIO(png_data))

        # Ensure the image is in RGB mode (BMP does not support transparency)
        if png_image.mode in ("RGBA", "LA"):
            png_image = png_image.convert("RGB")

        # Save the image as BMP to a byte stream
        bmp_stream = io.BytesIO()
        png_image.save(bmp_stream, format="BMP")

        # Get the BMP data as bytes
        bmp_data = bmp_stream.getvalue()

        return bmp_data
    except Exception as e:
        print(f"Error converting PNG to BMP: {e}")
        return None

encoded = uuencode_data(convert_png_to_bmp(base64.b64decode(result['images'][0])))

print(encoded)



