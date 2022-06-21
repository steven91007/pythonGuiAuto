import cv2
import os
import imageio

try:
    from PIL import Image
except ImportError:
    import Image


def recognize_structure(img, save_path, save_image=False):

    img_height, img_width = img.shape[:2]
    img = img[:img_height//6, :]
    gif_list = []
    img_original = img.copy()
    img_draw_box = img.copy()

    # print(img.shape)
    img = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    img_height, img_width = img.shape
    # print("img_height", img_height, "img_width", img_width)

    # thresholding the image to a binary image
    thresh, img_bin = cv2.threshold(img, 127, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU)
    img_bin = 255 - img_bin
    # img_bin = cv2.adaptiveThreshold(img, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 5, 5)
    # inverting the image

    # Plotting the image to see the output

    # countcol(width) of kernel as 100th of total width
    # kernel_len = np.array(img).shape[1] // 100
    kernel_len_ver = img_height // 50
    kernel_len_hor = img_width // 50
    # Defining a vertical kernel to detect all vertical lines of image
    ver_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, kernel_len_ver))  # shape (kernel_len, 1) inverted! xD
    # print("ver", ver_kernel)
    # print(ver_kernel.shape)

    # Defining a horizontal kernel to detect all horizontal lines of image
    hor_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (kernel_len_hor, 1))  # shape (1,kernel_ken) xD
    # print("hor", hor_kernel)
    # print(hor_kernel.shape)

    # A kernel of 2x2
    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (2, 2))
    # print(kernel)
    # print(kernel.shape)

    # Use vertical kernel to detect and save the vertical lines in a jpg
    image_1 = cv2.erode(img_bin, ver_kernel, iterations=3)
    vertical_lines = cv2.dilate(image_1, ver_kernel, iterations=4)

    # Plot the generated image

    # Use horizontal kernel to detect and save the horizontal lines in a jpg
    image_2 = cv2.erode(img_bin, hor_kernel, iterations=3)
    horizontal_lines = cv2.dilate(image_2, hor_kernel, iterations=4)

    # Plot the generated image
    # Combine horizontal and vertical lines in a new third image, with both having same weight.
    img_vh = cv2.addWeighted(vertical_lines, 0.5, horizontal_lines, 0.5, 0.0)

    # Eroding and thesholding the imagen
    img_vh = cv2.erode(~img_vh, kernel, iterations=2)

    thresh, img_vh = cv2.threshold(img_vh, 128, 255, cv2.THRESH_BINARY)

    bitxor = cv2.bitwise_xor(img, img_vh)
    bitnot = cv2.bitwise_not(bitxor)

    # Plotting the generated image

    # Detect contours for following box detectio
    contours, hierarchy = cv2.findContours(img_vh, cv2.RETR_TREE, cv2.CHAIN_APPROX_SIMPLE)

    # print(len(contours))
    # print(contours[0])
    # print(len(contours[0]))
    # print(cv2.boundingRect(contours[0]))

    def sort_contours(cnts, method="left-to-right"):
        # initialize the reverse flag and sort index
        reverse = False
        i = 0
        # handle if we need to sort in reverse
        if method == "right-to-left" or method == "bottom-to-top":
            reverse = True
        # handle if we are sorting against the y-coordinate rather than
        # the x-coordinate of the bounding box
        if method == "top-to-bottom" or method == "bottom-to-top":
            i = 1
        # construct the list of bounding boxes and sort them from top to
        # bottom
        bounding_boxes = [cv2.boundingRect(c) for c in cnts]
        (cnts, bounding_boxes) = zip(*sorted(zip(cnts, bounding_boxes),
                                             key=lambda b: b[1][i], reverse=reverse))
        # return the list of sorted contours and bounding boxes
        return cnts, bounding_boxes

    # Sort all the contours by top to bottom.
    contours, boundingBoxes = sort_contours(contours, method="top-to-bottom")

    # Create list box to store all boxes in
    box = []
    # Get position (x,y), width and height for every contour and show the contour on image
    # print("lencontours", len(contours))
    # print('-' * 20)
    # print(save_path)
    height, width, _ = img_original.shape
    flag = False
    for count, c in enumerate(contours):
        x, y, w, h = cv2.boundingRect(c)
        if w < 0.9 * img_width and h < 0.9 * img_height and w / h > 15 and w > h:
            # print('area:', w * h, 'w/h ratio:', w / h)
            radio = w / width
            if radio > 0.5:
                cv2.rectangle(img_draw_box, (x, y), (x + w, y + h), (0, 0, 255), 2)
                flag = True
                box.append([x, y, w, h])
    if save_image:

        if not os.path.exists(save_path):
            os.makedirs(save_path)

        path = os.path.join(save_path, 'original.png')
        gif_list.append(path)
        cv2.imwrite(path, img_original)
        path = os.path.join(save_path, 'gray.png')
        gif_list.append(path)
        cv2.imwrite(path, img)

        path = os.path.join(save_path, 'threshold.png')
        gif_list.append(path)
        cv2.imwrite(path, img_bin)
        path = os.path.join(save_path, 'vertical_lines.png')
        gif_list.append(path)
        cv2.imwrite(path, vertical_lines)
        path = os.path.join(save_path, 'horizontal_lines.png')
        gif_list.append(path)
        cv2.imwrite(path, horizontal_lines)
        path = os.path.join(save_path, 'img_vh.png')
        gif_list.append(path)
        cv2.imwrite(path, img_vh)
        path = os.path.join(save_path, 'img_vh_erode.png')
        gif_list.append(path)
        cv2.imwrite(path, img_vh)
        path = os.path.join(save_path, 'img_vh_thresh.png')
        gif_list.append(path)
        cv2.imwrite(path, img_vh)
        path = os.path.join(save_path, 'bitxor.png')
        gif_list.append(path)
        cv2.imwrite(path, bitxor)
        path = os.path.join(save_path, 'bitnot.png')
        gif_list.append(path)
        cv2.imwrite(path, bitnot)
        path = os.path.join(save_path, 'result.png')
        cv2.imwrite(path, img_draw_box)

        path = os.path.join(save_path, 'process.gif')
        images = []
        for i in gif_list:
            images.append(imageio.imread(i))

        # 存為gif動圖格式

        imageio.mimsave(path, images, fps=0.8)


    return _, flag


if __name__ == '__main__':
    img = cv2.imread(r'C:\Users\amo.cy.hsu\PycharmProjects\OCR\c.png')
    structure, _ = recognize_structure(img, save_path=r'C:\Users\amo.cy.hsu\PycharmProjects\OCR\test', save_image=True)
    [print(r) for r in structure]
