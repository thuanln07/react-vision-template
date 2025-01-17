import cv2
import numpy as np
import time
from pypylon import pylon
from tkinter import *
from PIL import Image, ImageTk
import threading  
import queue
import os 
from openpyxl import Workbook
from openpyxl.styles import Alignment
from datetime import datetime 
from tkinter import filedialog 
from tkinter import messagebox
import glob
import serial

#Connect camera
class CameraBasler:
    def __init__(self):
        try:
            self.camera = pylon.InstantCamera(pylon.TlFactory.GetInstance().CreateFirstDevice())
            self.camera.StartGrabbing(pylon.GrabStrategy_LatestImageOnly)
            self.converter = pylon.ImageFormatConverter()
            self.converter.OutputPixelFormat = pylon.PixelType_BGR8packed
            self.converter.OutputBitAlignment = pylon.OutputBitAlignment_MsbAligned 
        except Exception as e:
            print(f"Error initializing camera: {e}")
            self.camera = None
    
    def capture_picture(self, scale=1.0):
        if self.camera and self.camera.IsGrabbing():
            try:
                grabResult = self.camera.RetrieveResult(5000, pylon.TimeoutHandling_ThrowException)
                if grabResult.GrabSucceeded():
                    image = self.converter.Convert(grabResult)  
                    img_original = image.GetArray()
                    # Thu nhỏ ảnh nếu scale khác 1.0
                    if scale != 1.0:
                        height, width = img_original.shape[:2]
                        new_height, new_width = int(height * scale), int(width * scale)
                        img_original = cv2.resize(img_original, (new_width, new_height))
                    img_display = cv2.resize(img_original, (855, 641))
                    grabResult.Release()
                    return img_display
            except Exception as e:
                print(f"Error capturing image: {e}")
        return None
    
#xoay ảnh
def rotate_image(image, angle):
    (h, w) = image.shape[:2]
    center = (w // 2, h // 2)
    matrix = cv2.getRotationMatrix2D(center, angle, 1.0)
    rotated = cv2.warpAffine(image, matrix, (w, h), flags=cv2.INTER_LINEAR, borderMode=cv2.BORDER_CONSTANT, borderValue=0)
    return rotated
def combine_match_results(result_ccoeff, result_sqdiff, result_ccorr, weights=(0.5, 0.3, 0.2)):
    # Normalize SQDIFF
    result_sqdiff_normalized = 1 - result_sqdiff
    # Combine results
    combined_result = (weights[0] * result_ccoeff +
                       weights[1] * result_sqdiff_normalized +
                       weights[2] * result_ccorr)
    return combined_result
#pyramid template matching
def pyramid_template_matching(image, template, num_levels=4, threshold=0.99, angle_step=180):
    try:
        image_pyramid = [image]
        for _ in range(num_levels - 1):
            image_pyramid.append(cv2.pyrDown(image_pyramid[-1]))
        image_pyramid.reverse()

        template_pyramid = [template]
        for _ in range(num_levels - 1):
            template_pyramid.append(cv2.pyrDown(template_pyramid[-1]))
        template_pyramid.reverse()

        coarse_results = []
        for level in range(num_levels):
            current_image = image_pyramid[level]
            current_template = template_pyramid[level]
            scale_factor = 2 ** (num_levels - level - 1)

            for angle in range(0, 360, angle_step if level < num_levels - 1 else angle_step // 2):
                rotated_template = rotate_image(current_template, angle)

                # Apply multiple matching methods
                result_ccoeff = cv2.matchTemplate(current_image, rotated_template, cv2.TM_CCOEFF_NORMED)
                result_sqdiff = cv2.matchTemplate(current_image, rotated_template, cv2.TM_SQDIFF_NORMED)
                result_ccorr = cv2.matchTemplate(current_image, rotated_template, cv2.TM_CCORR_NORMED)

                # Combine results
                # Combine results
                combined_result = combine_match_results(result_ccoeff, result_sqdiff, result_ccorr)

                # Thresholding and location extraction
                locations = np.where(combined_result >= (threshold - level * 0.05))
                for pt in zip(*locations[::-1]):
                    coarse_results.append((combined_result[pt[1], pt[0]], (pt[0] * scale_factor, pt[1] * scale_factor), angle))

        return max(coarse_results, key=lambda x: x[0]) if coarse_results else None
    except Exception as e:
        print(f"Error in template matching: {e}")
        return None
    
class Application:
    
    def __init__(self, root):
        self.root = root
        self.root.title("VISION")
        self.root.geometry("1250x740")
        self.root.config(bg="skyblue")
        # Khởi tạo camera Basler
        self.basler_camera = CameraBasler()
        # Biến trạng thái
        self.running = False
        self.frame = None
        self.roi = None  # Khởi tạo vùng ROI

        #chia luồng chạy
        self.matching_thread = None
        self.result_queue = queue.Queue() 
        self.folder_path = None  # Đường dẫn thư mục chứa ảnh
        self.matching_timer_thread = None  # Luồng riêng cho tính năng matching sau 5 giây

        # Tải nhiều template
        self.template_files = [
                                 'image/template/TEM OK.png',
                                 'image/template/TEM NG.png',
                                 'image/template/TEM MISSING.png',
                               ]
        self.templates = [255 - cv2.imread(template_file, cv2.IMREAD_GRAYSCALE) for template_file in self.template_files]
        # Tạo giao diện
        self.create_ui()
        self.ser = serial.Serial('COM8',9600)
    def create_ui(self):
        main_frame = Frame(self.root, bg="skyblue")
        main_frame.pack(fill=BOTH, expand=True)
        # Tạo khung chứa các nút chức năng
        control_frame = Frame(main_frame, bg="#004d40", bd=2, relief=RIDGE, height=60)
        control_frame.pack(side=TOP, fill=X, pady=0)
        control_frame.pack_propagate(False)
        self.start_button = Button(control_frame, text="Start", width=15, command=self.start_matching, bg="#4caf50", fg="white", font=("Arial", 10, "bold"))
        self.start_button.pack(side=LEFT, padx=25, pady=10)
        self.stop_button = Button(control_frame, text="Stop", width=15, command=self.stop_matching, bg="#f44336", fg="white", font=("Arial", 10, "bold"))
        self.stop_button.pack(side=LEFT, padx=15, pady=10)
        self.capture_and_match_button = Button(control_frame, text="Capture", width=15, command=self.capture_and_match, bg="#2196f3", fg="white", font=("Arial", 10, "bold"))
        self.capture_and_match_button.pack(side=LEFT, padx=15, pady=10)
        self.delete_button = Button(control_frame, text="Delete", width=15, command=self.delete_info, bg="#ff9800", fg="white", font=("Arial", 10, "bold"))
        self.delete_button.pack(side=RIGHT, padx=28, pady=10)
        content_frame = Frame(main_frame, bg="skyblue")
        content_frame.pack(fill=BOTH, expand=True, padx=10, pady=10)
        image_frame = Frame(content_frame, bg="#eeeeee", bd=2, relief=RIDGE, width=900, height=650)
        image_frame.grid(row=0, column=0, padx=15, pady=10)
        image_frame.grid_propagate(False)
        self.canvas = Canvas(image_frame, width=855, height=640, bg="#eeeeee")
        self.canvas.pack(fill=BOTH, expand=True)
        info_frame = Frame(content_frame, bg="white", bd=2, relief=RIDGE, width=350, height=650)
        info_frame.grid(row=0, column=1, padx=15, pady=10, sticky=N)
        info_frame.grid_propagate(False)
        Label(info_frame, text="INFO CHECK TEM", font=("Arial", 14, "bold"), bg="white", fg="darkblue").pack(pady=15)
        label_width = 30
        self.result_label = Label(info_frame, text="Check: ", font=("Arial", 12), bg="white", fg="black", anchor=W, width=label_width, height=2)
        self.result_label.pack(fill=X, padx=10, pady=5)
        self.score_label = Label(info_frame, text="Score: ", font=("Arial", 12), bg="white", fg="black", anchor=W, width=label_width, height=2)
        self.score_label.pack(fill=X, padx=10, pady=5)
        self.location_label = Label(info_frame, text="Location: ", font=("Arial", 12), bg="white", fg="black", anchor=W, width=label_width, height=2)
        self.location_label.pack(fill=X, padx=10, pady=5)
        self.angle_label = Label(info_frame, text="Angle: ", font=("Arial", 12), bg="white", fg="black", anchor=W, width=label_width, height=2)
        self.angle_label.pack(fill=X, padx=10, pady=5)
        self.time_label = Label(info_frame, text="Time Matching: ", font=("Arial", 12), bg="white", fg="black", anchor=W, width=label_width, height=2)
        self.time_label.pack(fill=X, padx=10, pady=5)
        self.status_label = Label(info_frame, text="Status: Not yet started", font=("Arial", 12), bg="white", fg="black", anchor=W, width=label_width, height=2)
        self.status_label.pack(fill=X, padx=10, pady=5)
        self.load_image_button = Button(control_frame, text="Load Image", width=15, command=self.load_and_match_image, bg="#9c27b0", fg="white", font=("Arial", 10, "bold"))
        self.load_image_button.pack(side=LEFT, padx=15, pady=10)
        

    #Load image matching
    def load_and_match_image(self):
        """Chọn một file ảnh từ máy và thực hiện matching."""
        file_path = filedialog.askopenfilename(
            title="Select an Image File",
            filetypes=[("Image Files", "*.png;*.jpg;*.jpeg;*.bmp;*.tiff")]
        ) 
        if file_path:
            try:
                # Đọc ảnh từ file
                image = cv2.imread(file_path)
                if image is None:
                    print(f"Error: Cannot read the image file {file_path}")
                    return
                # Preprocess và thực hiện matching
                self.status_label.config(text="Status: Matching with image", fg="orange")
                start_time = time.time()
                # Crop theo ROI nếu có
                if self.roi:
                    x, y, w, h = self.roi
                    image = image[y:y+h, x:x+w]
                # processed_image = preprocess_image(image)  # Xử lý ảnh
                # self.display_processed_image(processed_image)  # Hiển thị ảnh đã xử lý
                image_gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
                image_gray = 255 - image_gray
                # cv2.imshow("")
                best_match = None
                best_template_name = ""
                for template, template_file in zip(self.templates, self.template_files):
                    match = pyramid_template_matching(image_gray, template, num_levels=2, threshold=0.7)
                    if match and (not best_match or match[0] > best_match[0]):
                        best_match = match
                        best_template_name = os.path.basename(template_file)
                end_time = time.time()
                matching_time = end_time - start_time
                # Cập nhật giao diện
                self.update_ui(image, best_match, best_template_name, matching_time)
            
            except Exception as e:
                print(f"Error loading and matching image: {e}")
                self.status_label.config(text="Error loading image.", fg="red")

    def display_processed_image(self, image):
        """Hiển thị ảnh xử lý lên canvas."""
        image_rgb = cv2.cvtColor(image, cv2.COLOR_GRAY2RGB)
        image_pil = Image.fromarray(image_rgb)
        image_tk = ImageTk.PhotoImage(image=image_pil)
        self.canvas.create_image(0, 0, anchor=NW, image=image_tk)
        self.canvas.image = image_tk

    def capture_and_match(self):
        frame = self.basler_camera.capture_picture(scale=1.0)
        self.status_label.config(text="Status: Matching", fg="orange")
        if frame is not None:
            start_time = time.time()
            # processed_frame = preprocess_image(frame)  # Xử lý ảnh trước khi matching
            # self.display_processed_image(processed_frame)  # Hiển thị ảnh đã xử lý
            image_gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
            image_gray = 255 - image_gray
            best_match = None
            best_template_name = ""
            for template, template_file in zip(self.templates, self.template_files):
                match = pyramid_template_matching(image_gray, template, num_levels=2, threshold=0.7)
                if match and (not best_match or match[0] > best_match[0]):
                    best_match = match
                    best_template_name = os.path.basename(template_file)
            end_time = time.time()
            matching_time = end_time - start_time
            self.update_ui(frame, best_match, best_template_name, matching_time)

    def display_processed_image(self, image):
        # Chuyển đổi ảnh xử lý thành định dạng có thể hiển thị trong Tkinter
        image_rgb = cv2.cvtColor(image, cv2.COLOR_GRAY2RGB)
        image_pil = Image.fromarray(image_rgb)
        image_tk = ImageTk.PhotoImage(image=image_pil)
        self.canvas.create_image(0, 0, anchor=NW, image=image_tk)
        self.canvas.image = image_tk

    def update_ui(self, frame, best_match, best_template_name, matching_time):
        self.result_label.config(text=f"CHECK: {best_template_name}")
        self.time_label.config(text=f"Time Matching: {matching_time:.4f}s")
        if best_match:
            score = best_match[0]
            location = best_match[1]
            angle = best_match[2]
            self.score_label.config(text=f"Score: {score:.4f}")
            self.location_label.config(text=f"Location: {location}")
            self.angle_label.config(text=f"Angle: {angle}")
        self.status_label.config(text="Status: Matching success", fg="green")

    def start_matching(self):
        self.running = True
        self.status_label.config(text="Status: Matching...",fg = "orange")
        if not self.matching_thread or not self.matching_thread.is_alive():
            self.matching_thread = threading.Thread(target=self.matching_loop, daemon=True)
            self.matching_thread.start()
        self.update_display()
        # self.running = False
        # self.status_label.config(text="Error initializing camera ")
    def stop_matching(self):
        self.running = False
        self.status_label.config(text="Status: Stop Matching...",fg = "red")
        # if self.matching_thread.is_alive():
        #     self.matching_thread.join()
        #     self.status_label.config(text="Trạng thái: Dừng matching...",fg = "yellow")
    def delete_info(self):
        self.result_label.config(text="CHECK ")
        self.score_label.config(text="Score: ")
        self.location_label.config(text="Location: ")
        self.angle_label.config(text="Angle: ")
        self.time_label.config(text="Time Matching: ")
        self.status_label.config(text="Status: Not yet started", fg="black")
        self.canvas.delete("all")

    def capture_and_match(self):
        threading.Thread(target=self.process_capture_and_match, daemon=True).start()

    def process_capture_and_match(self):
        frame = self.basler_camera.capture_picture(scale=1.0)
        if self.roi:
            x, y, w, h = self.roi
            frame = frame[y:y+h, x:x+w]  # Cắt ảnh theo ROI
        self.status_label.config(text="Status: Matching...", fg="orange")
        if frame is not None:
            start_time = time.time()
            image_gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
            image_gray = 255 - image_gray
            best_match = None
            best_template_name = ""
            for template, template_file in zip(self.templates, self.template_files):
                match = pyramid_template_matching(image_gray, template, num_levels=2, threshold=0.81)
                if match and (not best_match or match[0] > best_match[0]):
                    best_match = match
                    best_template_name = os.path.basename(template_file)
            end_time = time.time()
            matching_time = end_time - start_time
            self.update_ui(frame, best_match, best_template_name, matching_time)
    def matching_loop(self):
        while self.running:
            frame = self.basler_camera.capture_picture(scale=1.0)
            if frame is not None:
                start_time = time.time()
                image_gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
                image_gray = 255 - image_gray
                best_match = None
                best_template_name = ""
                for template, template_file in zip(self.templates, self.template_files):
                    match = pyramid_template_matching(image_gray, template, num_levels=2, threshold=0.80)
                    if match and (not best_match or match[0] > best_match[0]):
                        best_match = match
                        best_template_name = os.path.basename(template_file)
                end_time = time.time()
                matching_time = end_time - start_time
                self.result_queue.put((frame, best_match, best_template_name, matching_time))
                time.sleep(0.1)
   
    def update_display(self):
        if not self.result_queue.empty():
            frame, best_match, best_template_name, matching_time = self.result_queue.get()
            self.update_ui(frame, best_match, best_template_name, matching_time)
        if self.running:
            self.root.after(10, self.update_display)
    def update_ui(self, frame, best_match, best_template_name, matching_time):
        self.canvas.delete("all")
        frame_rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
        frame_pil = Image.fromarray(frame_rgb)
        frame_tk = ImageTk.PhotoImage(image=frame_pil)
        self.canvas.create_image(0, 0, anchor=NW, image=frame_tk)
        self.canvas.image = frame_tk
        self.result_label.config(text=f"CHECK: {best_template_name}")
        self.time_label.config(text=f"Time Matching: {matching_time:.4f}s")

        if self.roi:
            x, y, w, h = self.roi
            cv2.rectangle(frame, (x, y), (x+w, y+h), (0, 255, 0), 2)
        if best_match:
            score, location, angle = best_match
            # Vẽ và cập nhật hình chữ nhật xoay
            frame_color = frame.copy()  # Tạo bản sao của khung hình gốc để vẽ
            template_name_only = os.path.splitext(os.path.basename(best_template_name))[0]
            h, w = cv2.imread(f"image/template/{best_template_name}", cv2.IMREAD_GRAYSCALE).shape

            # Tính toán tọa độ hình chữ nhật xoay
            box = cv2.boxPoints(((location[0] + w / 2, location[1] + h / 2), (w, h), -angle))
            box = np.intp(box)
            
            # Vẽ hình chữ nhật xoay lên khung hình
            cv2.drawContours(frame_color, [box], 0, (0, 0, 255), 2)  

            # Chọn màu chữ dựa vào template được matching
            if template_name_only == "TEM OK":
                text_color = (0, 255, 0)  # Màu xanh cho template 1
            elif template_name_only == "TEM NG":
                text_color = (0, 0, 255)  # Màu đỏ cho template 2
            elif template_name_only == "TEM MiSSING":
                text_color = (0, 0, 255)  # Màu đỏ cho template 2
            else:
                text_color = (0, 0, 255)  # Mặc định là màu trắng nếu không xác định

            # Thêm thông tin template lên khung hình
            cv2.putText(frame_color, f"{template_name_only}", 
                        (int(location[0]), int(location[1]) - 10),
                        cv2.FONT_HERSHEY_SIMPLEX, 0.8, text_color, 2)  
              
            
            # Hiển thị khung hình có chứa hình chữ nhật lên canvas
            frame_rgb = cv2.cvtColor(frame_color, cv2.COLOR_BGR2RGB)
            frame_pil = Image.fromarray(frame_rgb)
            frame_tk = ImageTk.PhotoImage(image=frame_pil)
            self.canvas.create_image(0, 0, anchor=NW, image=frame_tk)
            self.canvas.image = frame_tk
            # Cập nhật các thông tin hiển thị
            self.result_label.config(text=f"CHECK: {template_name_only}")
            self.score_label.config(text=f"Score: {score * 100:.2f}%")
            self.location_label.config(text=f"Location: (X: {int(location[0])}, Y: {int(location[1])})")
            self.angle_label.config(text=f"Angle: {angle}°")
            self.time_label.config(text=f"Time Matching: {matching_time:.2f}s")
            self.status_label.config(text="Status: Matching success", fg="green")
            # Ghi vào file Excel
            #self.save_to_excel(best_template_name, score, location, angle, matching_time)
            print(f"CHECK: {template_name_only}")
            print(f"Score: {score * 100:.2f}%")
            print(f"Location: (X: {int(location[0])}, Y: {int(location[1])})")
            print(f"Angle: {angle}°")
            print(f"Time Matching: {matching_time:.2f}s")
            print("----------------------------------------------------------------------")
            #send serial command
            if template_name_only == "TEM OK":
                self.ser.write(b'1')
                print("Send serial command: 1")
            elif template_name_only == "TEM NG":
                self.ser.write(b'2')
                print("Send serial command: 2")
            elif template_name_only == "TEM MISSING":
                self.ser.write(b'3')
                print("Send serial command: 3")
            else:
                self.ser.write(b'4')
                print("Send serial command: 4 ")
        else:
            self.result_label.config(text="CHECK: Not determined")
            self.score_label.config(text="Score: Not determined")
            self.location_label.config(text="Location: Not determined")
            self.angle_label.config(text="Angle: Not determined")
            self.time_label.config(text=f"Time Matching: {matching_time:.2f}s")
            self.status_label.config(text="Status:Not determined", fg="red")
            print("No sample found")
            print(f"Time Matching: {matching_time:.2f}s")
            print("----------------------------------------------------------------------")

if __name__ == "__main__":
     root = Tk()
     app = Application(root)
     root.mainloop()