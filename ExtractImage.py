import fitz
import io
from PIL import Image


file = "C:\\Users\\PRAJWAL\\PycharmProjects\\WebScapping\\excel\\iMAGES3.pdf"
pdf_file = fitz.open(file)

imglist=[]
for page_index in range(len(pdf_file)):
    page = pdf_file[page_index]
    image_list = page.get_images()
    if image_list:
        print(f"[+] Found a total of {len(image_list)} images in page {page_index}")
    else:
        print("[!] No images found on page", page_index)
    for image_index, img in enumerate(image_list, start=1):
        xref = img[0]
        base_image = pdf_file.extract_image(xref)
        image_bytes = base_image["image"]

        # get the image extension
        image_ext = base_image["ext"]
        image = Image.open(io.BytesIO(image_bytes))
        image.save(open(f"image{page_index + 1}_{image_index}.{image_ext}", "wb"))
        im = image.convert('RGB')
        imglist.append(im)
        im.save(r"C:\\Users\\PRAJWAL\\PycharmProjects\\WebScapping\\excel\\StartUpsImage.pdf",save_all=True, append_images=imglist)


# file= "C:\\Users\\PRAJWAL\\PycharmProjects\\WebScapping\\excel\\Object.pdf"
# pdf_file=fitz.open(file)
#
# # iterate over PDF pages
# for page_index in range(len(pdf_file)):
#     # get the page itself
#     page = pdf_file[page_index]
#     # get image list
#     image_list = page.get_images()
#     # printing number of images found in this page
#     if image_list:
#         print(f"[+] Found a total of {len(image_list)} images in page {page_index}")
#     else:
#         print("[!] No images found on page", page_index)
#     for image_index, img in enumerate(image_list, start=1):
#         # get the XREF of the image
#         xref = img[0]
#         # extract the image bytes
#         base_image = pdf_file.extract_image(xref)
#         image_bytes = base_image["image"]
#         # get the image extension
#         image_ext = base_image["ext"]
#         # load it to PIL
#         image = Image.open(io.BytesIO(image_bytes))
#         # save it to local disk
#         image.save(open(f"image{page_index+1}_{image_index}.{image_ext}", "wb"))


