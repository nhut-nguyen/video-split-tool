from moviepy.video.io.VideoFileClip import VideoFileClip
import os
import pandas as pd

excel_file = "video_split.xlsx"
input_folder = "videos"
output_folder = "videos_output"

data = pd.read_excel(excel_file)

for index, row in data.iterrows():
    input_video = os.path.join(input_folder, row['file_name'])  # Đường dẫn video gốc
    start_time = float(row['from'])  # Thời gian bắt đầu
    end_time = float(row['to'])  # Thời gian kết thúc
    label = row['label']  # Nhãn
    video_index = row['index']  # Thứ tự

    # Tạo tên file đầu ra dựa trên nhãn
    output_video = os.path.join(output_folder, f"{label}_{video_index}.mp4")

    try:
        # Đọc video và cắt
        with VideoFileClip(input_video) as video:
            cut_video = video.subclipped(start_time, end_time)  # Cắt video
            cut_video = cut_video.without_audio()  # Loại bỏ âm thanh
            cut_video.write_videofile(output_video, codec="libx264")

        print(f"Video {video_index}. {row['file_name']} đã được xử lý và lưu tại {output_video}")
    except Exception as e:
        print(f"Lỗi khi xử lý video {video_index}. {row['file_name']}: {e}")