import cv2
import os
import numpy as np
from pptx import Presentation
from pptx.util import Inches

# ================= 配置区 =================
INTERVAL_SECONDS = 1      # 采样频率
SENSITIVITY_FACTOR = 1  # 突变倍数：当前变化是平均变化的3倍时，视为翻页
# ==========================================

def video_to_ppt(video_path):
    cap = cv2.VideoCapture(video_path)
    prs = Presentation()
    fps = cap.get(cv2.CAP_PROP_FPS) or 30
    frame_interval = int(fps * INTERVAL_SECONDS)
    
    last_gray = None 
    pending_frame = None   
    diff_history = []      # 存储近期像素变化量的历史，用于计算平均值
    
    slide_count = 0
    count = 0

    print(f"\n▶ 正在运行“动态突变检测”: {video_path}")
    while cap.isOpened():
        ret, frame = cap.read()
        if not ret: 
            if pending_frame is not None:
                add_slide(prs, pending_frame)
            break
        
        if count % frame_interval == 0:
            gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
            gray_blur = cv2.GaussianBlur(gray, (21, 21), 0)
            
            if last_gray is not None:
                # 1. 计算当前像素变化
                delta = cv2.absdiff(last_gray, gray_blur)
                thresh = cv2.threshold(delta, 25, 255, cv2.THRESH_BINARY)[1]
                current_diff = cv2.countNonZero(thresh)
                
                # 2. 计算历史平均变化（取最近5次采样）
                avg_diff = np.mean(diff_history) if len(diff_history) > 0 else current_diff
                
                # 打印数据方便观察：当前变化量 vs 历史均值
                print(f"  当前变化: {current_diff} | 均值: {avg_diff:.1f}  ", end="\r")

                # 3. 核心判断：如果当前变化发生了“异常突变”
                if len(diff_history) > 3 and current_diff > avg_diff * SENSITIVITY_FACTOR:
                    # 只有当变化量突然飙升，说明翻页了
                    if pending_frame is not None:
                        add_slide(prs, pending_frame) # 保存翻页前内容最满的一张
                        slide_count += 1
                        print(f"\n[!] 捕捉到突变！(变化量是均值的 {current_diff/avg_diff:.1f} 倍) 已保存旧页。")
                        diff_history = [] # 翻页后重置历史，适应新页面的节奏
                
                # 4. 更新逻辑
                if current_diff > 50: # 忽略微小的噪点波动
                    diff_history.append(current_diff)
                    if len(diff_history) > 10: diff_history.pop(0)
                
                pending_frame = frame.copy() # 始终更新暂存帧，保证它是最新的
                last_gray = gray_blur
            else:
                last_gray = gray_blur
                pending_frame = frame.copy()
            
        count += 1
    
    output_name = video_path.rsplit('.', 1)[0] + "_dynamic.pptx"
    prs.save(output_name)
    cap.release()
    print(f"\n✅ 完成！共生成 {slide_count} 页课件。")

def add_slide(prs, frame):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    cv2.imwrite("temp.jpg", frame)
    slide.shapes.add_picture("temp.jpg", 0, 0, width=Inches(10))
    if os.path.exists("temp.jpg"): os.remove("temp.jpg")

for file in os.listdir("."):
    if file.endswith(".mp4"):
        video_to_ppt(file)