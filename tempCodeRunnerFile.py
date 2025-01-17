
# def rotate_image(image, angle):
#     (h, w) = image.shape[:2]
#     center = (w // 2, h // 2)
#     matrix = cv2.getRotationMatrix2D(center, angle, 1.0)
#     rotated = cv2.warpAffine(image, matrix, (w, h), flags=cv2.INTER_LINEAR, borderMode=cv2.BORDER_CONSTANT, borderValue=0)
#     return rotated

# def pyramid_template_matching(image, template, num_levels=4, threshold=0.99, angle_step=50):
#     try:
#         image_pyramid = [image]
#         for _ in range(num_levels - 1):
#             image_pyramid.append(cv2.pyrDown(image_pyramid[-1]))
#         image_pyramid.reverse()

#         template_pyramid = [template]
#         for _ in range(num_levels - 1):
#             template_pyramid.append(cv2.pyrDown(template_pyramid[-1]))
#         template_pyramid.reverse()

#         coarse_results = []
#         for level in range(num_levels):
#             current_image = image_pyramid[level]
#             current_template = template_pyramid[level]
#             scale_factor = 2 ** (num_levels - level - 1)

#             for angle in range(0, 360, angle_step if level < num_levels - 1 else angle_step // 2):
#                 rotated_template = rotate_image(current_template, angle)

#                 # Apply multiple matching methods
#                 result_ccoeff = cv2.matchTemplate(current_image, rotated_template, cv2.TM_CCOEFF_NORMED)
#                 result_sqdiff = cv2.matchTemplate(current_image, rotated_template, cv2.TM_SQDIFF_NORMED)

#                 # Normalize the results (since TM_SQDIFF_NORMED prefers lower values, invert it)
#                 result_combined = (result_ccoeff + (1 - result_sqdiff)) / 2

#                 # Thresholding and location extraction
#                 locations = np.where(result_combined >= (threshold - level * 0.05))
#                 for pt in zip(*locations[::-1]):
#                     coarse_results.append((result_combined[pt[1], pt[0]], (pt[0] * scale_factor, pt[1] * scale_factor), angle))

#         return max(coarse_results, key=lambda x: x[0]) if coarse_results else None
#     except Exception as e:
#         print(f"Error in template matching: {e}")
#         return None