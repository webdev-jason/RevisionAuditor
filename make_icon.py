from PIL import Image

# 1. Change this to the exact name of the image you downloaded from me
input_image_path = "RevisionAuditor.png" 

# 2. Open the image
img = Image.open(input_image_path)

# 3. Define the sizes Windows needs for the icon to scale properly
icon_sizes = [(16, 16), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)]

# 4. Save it as a multi-size .ico file
img.save("icon.ico", format="ICO", sizes=icon_sizes)

print("Success! Your multi-size icon.ico file has been created.")