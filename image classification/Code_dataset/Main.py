import face_recognition
import cv2
import numpy as np
import pandas as pd
import os
import time
from datetime import datetime
from email.message import EmailMessage
import ssl
import smtplib
import json
import glob

# Load configuration from config.json
def load_config():
    with open("config.json", "r") as config_file:
        config = json.load(config_file)
    return config

config = load_config()

def get_time_period():
    """
    Determine the time period based on current time and config.
    """
    current_hour = datetime.now().hour
    for period, hours in config["time_periods"].items():
        if hours["start"] <= current_hour < hours["end"]:
            return period
    return "default"

# Generate Excel file name with current date and time period
current_date = datetime.now().strftime("%d-%m-%Y")
time_period = get_time_period()
EXCEL_FILE_PATH = f"{current_date}-{time_period}.xlsx"

def load_known_faces():
    """
    Load and encode known face images from config.
    Returns the known face encodings and corresponding names.
    """
    known_face_encodings = []
    known_face_names = []

    for image_path, name in config["known_faces"]:
        try:
            image = face_recognition.load_image_file(image_path)
            encoding = face_recognition.face_encodings(image)[0]
            known_face_encodings.append(encoding)
            known_face_names.append(name)
        except IndexError:
            print(f"Warning: No face found in image {image_path}. Skipping this image.")

    return known_face_encodings, known_face_names

def save_recognized_faces_status(status_dict):
    """
    Save the recognized face names and their statuses (Present/Absent) to the Excel file.
    """
    data = []
    for name, status in status_dict.items():
        data.append([name, status])

    df = pd.DataFrame(data, columns=['Name', 'Status'])
    df.to_excel(EXCEL_FILE_PATH, index=False)
    print("Time's up")
    print(f"Attendance saved to {EXCEL_FILE_PATH}")

def send_email_with_excel():
    """
    Send an email with the attached Excel file after saving the attendance.
    """
    email_sender = config["email"]["sender"]
    email_password = config["email"]["password"]
    email_receiver = config["email"]["receiver"]

    subject = "Today's attendance Sir/Mam:"
    body = """
    Reporting Today's Attendance:
    """

    em = EmailMessage()
    em['From'] = email_sender
    em['To'] = email_receiver
    em['Subject'] = subject
    em.set_content(body)

    if os.path.exists(EXCEL_FILE_PATH):
        with open(EXCEL_FILE_PATH, 'rb') as f:
            file_data = f.read()
            file_name = os.path.basename(EXCEL_FILE_PATH)
            em.add_attachment(file_data, maintype='application', subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet', filename=file_name)
    else:
        print("File not found")

    context = ssl.create_default_context()

    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as smtp:
            smtp.login(email_sender, email_password)
            smtp.sendmail(email_sender, email_receiver, em.as_string())
        print(f"Email sent successfully to {email_receiver} with the attendance Excel file attached!")
    except Exception as e:
        print(f"Failed to send email: {e}")

def send_parent_emails(attendance_status):
    """
    Send individual emails to parents about their child's attendance status if the student is absent.
    """
    email_sender = config["email"]["sender"]
    email_password = config["email"]["password"]
    parent_emails = config["parent_emails"]
    context = ssl.create_default_context()
    current_date = datetime.now().strftime("%B %d, %Y")

    for student_name, status in attendance_status.items():
        if student_name in parent_emails and "Absent" in status:
            parent_email = parent_emails[student_name]

            em = EmailMessage()
            em['From'] = email_sender
            em['To'] = parent_email
            em['Subject'] = f"Attendance Report for {student_name} - {current_date}"

            body = f"""
            Dear Parent,

            This is to inform you that {student_name} was absent for college today, {current_date}.

            Best regards,
            College Administration
            """

            em.set_content(body)

            try:
                with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as smtp:
                    smtp.login(email_sender, email_password)
                    smtp.sendmail(email_sender, parent_email, em.as_string())
                print(f"Attendance notification sent to parent of {student_name}")
            except Exception as e:
                print(f"Failed to send email for {student_name}: {e}")

def check_morning_attendance(current_attendance):
    """
    Check morning attendance file to find students who were present in morning but absent in afternoon.
    Returns list of such students.
    """
    morning_file_pattern = f"{current_date}-morning.xlsx"
    morning_files = glob.glob(morning_file_pattern)
    
    if not morning_files:
        print("No morning attendance file found.")
        return []
    
    morning_file = morning_files[0]
    try:
        morning_df = pd.read_excel(morning_file)
        morning_present = morning_df[morning_df['Status'].str.contains('Present')]['Name'].tolist()
        
        afternoon_absent = [name for name, status in current_attendance.items() if "Absent" in status]
        
        # Students present in morning but absent in afternoon
        problematic_students = list(set(morning_present) & set(afternoon_absent))
        
        return problematic_students
    except Exception as e:
        print(f"Error reading morning attendance file: {e}")
        return []

def notify_problematic_attendance(problematic_students):
    """
    Send notifications about students who were present in morning but absent in afternoon.
    """
    if not problematic_students:
        return
    
    email_sender = config["email"]["sender"]
    email_password = config["email"]["password"]
    email_receiver = config["email"]["receiver"]
    parent_emails = config["parent_emails"]
    context = ssl.create_default_context()
    current_date = datetime.now().strftime("%B %d, %Y")
    
    # Send to admin
    admin_subject = "Attendance Alert: Students Present in Morning but Absent in Afternoon"
    admin_body = f"""
    The following students were present in the morning session but are absent in the afternoon session on {current_date}:
    
    {', '.join(problematic_students)}
    
    Please take necessary action.
    """
    
    em = EmailMessage()
    em['From'] = email_sender
    em['To'] = email_receiver
    em['Subject'] = admin_subject
    em.set_content(admin_body)
    
    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as smtp:
            smtp.login(email_sender, email_password)
            smtp.sendmail(email_sender, email_receiver, em.as_string())
        print("Admin notification sent about problematic attendance")
    except Exception as e:
        print(f"Failed to send admin email: {e}")
    
    # Send to parents
    for student in problematic_students:
        if student in parent_emails:
            parent_email = parent_emails[student]
            
            em = EmailMessage()
            em['From'] = email_sender
            em['To'] = parent_email
            em['Subject'] = f"Attendance Alert for {student} - {current_date}"
            
            body = f"""
            Dear Parent,
            
            This is to inform you that {student} was present in the morning session 
            but is absent in the afternoon session today, {current_date}.
            
            Please contact the college administration if there's any discrepancy.
            
            Best regards,
            College Administration
            """
            
            em.set_content(body)
            
            try:
                with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as smtp:
                    smtp.login(email_sender, email_password)
                    smtp.sendmail(email_sender, parent_email, em.as_string())
                print(f"Problematic attendance notification sent to parent of {student}")
            except Exception as e:
                print(f"Failed to send email for {student}: {e}")

def recognize_faces_from_video():
    """
    Capture video from the webcam and recognize faces in real-time.
    """
    known_face_encodings, known_face_names = load_known_faces()
    attendance_status = {name: "Absent" for name in known_face_names}
    recognized_faces = {}
    video_capture = cv2.VideoCapture(0)
    start_time = time.time()
    time_limit = config["time_limit"]

    while True:
        current_time = time.time()
        ret, frame = video_capture.read()

        if not ret:
            print("Failed to capture video frame.")
            break

        rgb_frame = frame[:, :, ::-1]
        small_frame = cv2.resize(rgb_frame, (0, 0), fx=0.25, fy=0.25)
        face_locations = face_recognition.face_locations(small_frame)
        face_encodings = face_recognition.face_encodings(small_frame, face_locations)

        for (top, right, bottom, left), face_encoding in zip(face_locations, face_encodings):
            matches = face_recognition.compare_faces(known_face_encodings, face_encoding, tolerance=0.45)
            name = "Unknown"
            face_distances = face_recognition.face_distance(known_face_encodings, face_encoding)
            best_match_index = np.argmin(face_distances)

            if matches[best_match_index]:
                name = known_face_names[best_match_index]

            top *= 4
            right *= 4
            bottom *= 4
            left *= 4

            box_color = (0, 255, 0) if name != "Unknown" else (0, 0, 255)
            cv2.rectangle(frame, (left, top), (right, bottom), box_color, 2)
            font = cv2.FONT_HERSHEY_DUPLEX
            cv2.putText(frame, name, (left + 6, bottom - 6), font, 0.5, box_color, 1)

            if current_time - start_time <= time_limit:
                if name != "Unknown" and name not in recognized_faces:
                    recognized_faces[name] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    attendance_status[name] = f"Present on {recognized_faces[name]}"
                    print(f"{name} marked as Present at {recognized_faces[name]}")

        cv2.imshow('Face_Capture', frame)

        if cv2.waitKey(1) & 0xFF == ord('q') or current_time - start_time > time_limit + 5:
            break

    current_date_time = datetime.now().strftime("%Y-%m-%d")
    for name in attendance_status:
        if name not in recognized_faces:
            attendance_status[name] = f"Absent on {current_date_time}"

    save_recognized_faces_status(attendance_status)
    video_capture.release()
    cv2.destroyAllWindows()

    # Send the Excel file to the admin
    send_email_with_excel()

    # Check for students present in morning but absent in afternoon
    print(time_period)
    if time_period == "afternoon":
        problematic_students = check_morning_attendance(attendance_status)
        if problematic_students:
            notify_problematic_attendance(problematic_students)

    # Send individual emails to parents for morning absentees
    if time_period == "morning":
        print("Sending attendance notifications to parents...")
        send_parent_emails(attendance_status)
    else:
        print("Skipping parent notifications - not morning period")

if __name__ == "__main__":
    recognize_faces_from_video()
