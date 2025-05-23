from flask import Flask, render_template, request, send_file, flash, redirect, url_for
from googleapiclient.discovery import build
import pandas as pd
import io
from isodate import parse_duration
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
import os

app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET_KEY", "super_secret_key")  # Ortam değişkeninden al, fallback ile
API_KEY = os.environ.get("YOUTUBE_API_KEY")  # API anahtarını ortam değişkeninden al

if not API_KEY:
    raise ValueError("YOUTUBE_API_KEY ortam değişkeni tanımlı değil!")

def extract_channel_id(input_str):
    try:
        if "youtube.com/channel/" in input_str:
            return input_str.split("youtube.com/channel/")[1].split("/")[0]
        elif "youtube.com/@" in input_str:
            service = build("youtube", "v3", developerKey=API_KEY)
            username = input_str.split("youtube.com/@")[1].split("/")[0]
            response = service.search().list(q=username, type="channel", part="snippet", maxResults=1).execute()
            if response["items"]:
                return response["items"][0]["snippet"]["channelId"]
            else:
                raise ValueError("Kanal bulunamadı.")
        return input_str
    except Exception as e:
        raise ValueError(f"Kanal ID alınamadı: {str(e)}")

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        if "channel_input" in request.form:
            channel_input = request.form["channel_input"]
            try:
                channel_id = extract_channel_id(channel_input)
                flash("Kanal ID başarıyla bulundu!", "success")
                return render_template("index.html", channel_id=channel_id)
            except Exception as e:
                flash(str(e), "error")
                return redirect(url_for("index"))

        elif "channel_id" in request.form:
            channel_id = request.form["channel_id"]
            try:
                service = build("youtube", "v3", developerKey=API_KEY)
                uploads_id = service.channels().list(part="contentDetails", id=channel_id).execute()["items"][0]["contentDetails"]["relatedPlaylists"]["uploads"]

                videos = []
                next_page_token = None

                while True:
                    pl_request = service.playlistItems().list(
                        part="snippet",
                        playlistId=uploads_id,
                        maxResults=50,
                        pageToken=next_page_token
                    )
                    pl_response = pl_request.execute()

                    video_ids = [item["snippet"]["resourceId"]["videoId"] for item in pl_response["items"]]

                    if video_ids:
                        video_request = service.videos().list(
                            part="snippet,statistics,contentDetails",
                            id=",".join(video_ids)
                        )
                        video_response = video_request.execute()

                        for item in video_response["items"]:
                            duration = parse_duration(item["contentDetails"]["duration"]).total_seconds()
                            if duration > 60:  # Shorts'ları filtrele (60 saniyeden kısa videolar hariç)
                                video = {
                                    "Başlık": item["snippet"]["title"],
                                    "Yayın Tarihi": item["snippet"]["publishedAt"],
                                    "Video URL": f"https://www.youtube.com/watch?v={item['id']}",
                                    "İzlenme Sayısı": item["statistics"].get("viewCount", "0")
                                }
                                videos.append(video)

                    next_page_token = pl_response.get("nextPageToken")
                    if not next_page_token:
                        break

                if not videos:
                    flash("Kanalda 60 saniyeden uzun video bulunamadı.", "error")
                    return redirect(url_for("index"))

                # Pandas DataFrame oluştur
                df = pd.DataFrame(videos)

                # Excel dosyası oluştur
                wb = Workbook()
                ws = wb.active
                ws.title = "YouTube Videoları"

                # Başlıkları stil ile ekle
                headers = ["Başlık", "Yayın Tarihi", "Video URL", "İzlenme Sayısı"]
                for col_num, header in enumerate(headers, 1):
                    cell = ws.cell(row=1, column=col_num)
                    cell.value = header
                    cell.font = Font(bold=True)
                    cell.alignment = Alignment(horizontal="center")
                    cell.border = Border(left=Side(style="thin"), right=Side(style="thin"),
                                        top=Side(style="thin"), bottom=Side(style="thin"))

                # Verileri ekle
                for row in dataframe_to_rows(df, index=False, header=False):
                    ws.append(row)

                # Sütun genişliklerini ayarla
                for col in ws.columns:
                    max_length = 0
                    column = col[0].column_letter
                    for cell in col:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    adjusted_width = max_length + 2
                    ws.column_dimensions[column].width = adjusted_width

                # Excel dosyasını kaydet
                excel_file = io.BytesIO()
                wb.save(excel_file)
                excel_file.seek(0)

                return send_file(
                    excel_file,
                    mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    as_attachment=True,
                    download_name="youtube_videolari.xlsx"
                )

            except Exception as e:
                flash(f"Video bilgileri alınamadı: {str(e)}", "error")
                return redirect(url_for("index"))

    return render_template("index.html")

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5050)), debug=False)