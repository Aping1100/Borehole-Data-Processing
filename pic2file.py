import cv2, os
import numpy as np


def detect_and_transform(input_folder, output_folder):
    os.makedirs(output_folder, exist_ok=True)
    image_files = [f for f in os.listdir(input_folder) if f.endswith(('.png', '.jpg', '.jpeg'))]
    print_message=[]
    for image_file in image_files:
        # compose file path
        image_path = os.path.join(input_folder, image_file)

        # Read pic
        image = cv2.imread(image_path)
        hsv_image = cv2.cvtColor(image, cv2.COLOR_BGR2HSV)

        # Select color range
        lower_blue = np.array([90, 90, 90])
        upper_blue = np.array([100, 255, 255])
        # Mask
        mask = cv2.inRange(hsv_image, lower_blue, upper_blue)

        result = cv2.bitwise_and(image, image, mask=mask)
        # Convert to grayscale
        gray_result = cv2.cvtColor(result, cv2.COLOR_BGR2GRAY)
        # Gaussian blur
        blurred = cv2.GaussianBlur(gray_result, (5, 5), 0)
        # contour detect
        edges = cv2.Canny(blurred, 100, 200)  # adjust threshold

        # Dilation and erosion
        kernel = np.ones((5, 5), np.uint8)
        edges = cv2.dilate(edges, kernel, iterations=1)
        edges = cv2.erode(edges, kernel, iterations=1)

        # contour detect
        contours, _ = cv2.findContours(edges, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        contours = sorted(contours, key=cv2.contourArea, reverse=True)
        # Sort based on contour area
        largest_contour = max(contours, key=cv2.contourArea)

        # Approximate polygon
        epsilon = 0.02 * cv2.arcLength(largest_contour, True)
        approx = cv2.approxPolyDP(largest_contour, epsilon, True)
        cv2.drawContours(image, [approx], -1, (0, 255, 0), 2)

        # The number of vertices detected in the quadrilateral
        if len(approx) == 4:
            approx = approx.tolist()
            # print(approx)
            points = [approx[0][0], approx[1][0], approx[2][0], approx[3][0]]
            # print(points)
            points = np.array(points)

            # Find the centroid of the points
            center = np.mean(points, axis=0)

            # Calculate the polar angle of each point with respect to the centroid.
            angles = np.arctan2(points[:, 1] - center[1], points[:, 0] - center[0])

            # order
            sorted_indices = np.argsort(angles)
            sorted_points = points[sorted_indices]

            box = np.int0(sorted_points)
            box = np.array(box, dtype=np.float32)
            # transform
            for point in box:
                cv2.circle(image, tuple(point.astype(int)), 5, (0, 255, 0), -1)
            target_corners = np.array([[0, 0], [600, 0], [600, 200], [0, 200]], dtype=np.float32)

            M = cv2.getPerspectiveTransform(box, target_corners)
            transformed_image = cv2.warpPerspective(image, M, (600, 200))
            # save file
            output_path = os.path.join(output_folder, f"transformed_{image_file}")
            print("轉換完成", image_file)
            print_message.append(f"轉換完成 {image_file}")
            cv2.imwrite(output_path, transformed_image)


        else:
            print("未檢測到四邊形", image_file)
            print_message.append(f"未檢測到四邊形{image_file}")

    return print_message


