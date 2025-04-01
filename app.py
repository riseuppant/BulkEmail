from flask import Flask, render_template, request, redirect, url_for, flash
import os
import smtplib
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.secret_key = "your_secret_key"  # Needed for flash messages

# Configuration
UPLOAD_FOLDER = 'uploads'
IMAGES_FOLDER = 'images'  # Folder containing PNG images
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['IMAGES_FOLDER'] = IMAGES_FOLDER

# Email configuration
SMTP_SERVER = 'smtp.gmail.com'  # Change as needed
SMTP_PORT = 587
EMAIL_ADDRESS = 'hardikbatwal1505@gmail.com'  # Your email
EMAIL_PASSWORD = 'wzzh xztp zgcd ifvs'  # Your app password

# Create necessary directories if they don't exist
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(IMAGES_FOLDER, exist_ok=True)

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload_excel', methods=['POST'])
def upload_excel():
    if request.method == 'POST':
        # Check if an Excel file was uploaded
        if 'excel_file' not in request.files:
            flash('No Excel file uploaded')
            return redirect(request.url)
        
        excel_file = request.files['excel_file']
        
        # If user doesn't select a file
        if excel_file.filename == '':
            flash('No Excel file selected')
            return redirect(request.url)
        
        if excel_file and allowed_file(excel_file.filename):
            filename = secure_filename(excel_file.filename)
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            excel_file.save(filepath)
            
            # Process Excel file and send emails
            subject = request.form['subject']
            message_template = request.form['message']
            
            try:
                # Read Excel file
                df = pd.read_excel(filepath)
                
                # Check if required columns exist
                if 'email' not in df.columns or 'image_name' not in df.columns:
                    flash('Excel file must contain "email" and "image_name" columns')
                    os.remove(filepath)
                    return redirect(url_for('index'))
                
                successful = 0
                failed = 0
                failed_emails = []
                
                # Process each row
                for _, row in df.iterrows():
                    email = row['email']
                    image_name = row['image_name']
                    
                    # Check if image exists
                    image_path = os.path.join(app.config['IMAGES_FOLDER'], image_name)
                    if not os.path.exists(image_path):
                        failed += 1
                        failed_emails.append(f"{email} (Image not found: {image_name})")
                        continue
                    
                    # Create personalized message if placeholders exist
                    personalized_message = message_template
                    for column in df.columns:
                        if f"{{{column}}}" in message_template:
                            personalized_message = personalized_message.replace(f"{{{column}}}", str(row[column]))
                    
                    # Try to send the email
                    success = send_email_with_image(email, subject, personalized_message, image_path)
                    
                    if success:
                        successful += 1
                    else:
                        failed += 1
                        failed_emails.append(email)
                
                # Clean up the uploaded file
                os.remove(filepath)
                
                if failed == 0:
                    flash(f'Successfully sent {successful} emails!')
                else:
                    flash(f'Sent {successful} emails. Failed to send {failed} emails.')
                    for failed_email in failed_emails:
                        flash(f'Failed: {failed_email}')
                
            except Exception as e:
                flash(f'Error processing Excel file: {str(e)}')
                
            return redirect(url_for('index'))
        else:
            flash('Only Excel files (.xlsx, .xls) are allowed')
            return redirect(request.url)

def send_email_with_image(recipient, subject, message_body, image_path):
    try:
        # Create email message
        msg = MIMEMultipart()
        msg['From'] = EMAIL_ADDRESS
        msg['To'] = recipient
        msg['Subject'] = subject
        
        # Attach text part
        msg.attach(MIMEText(message_body, 'plain'))
        
        # Attach image
        with open(image_path, 'rb') as f:
            img_data = f.read()
            img = MIMEImage(img_data, name=os.path.basename(image_path))
            msg.attach(img)
        
        # Connect to SMTP server
        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        server.starttls()
        server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
        
        # Send email
        server.send_message(msg)
        server.quit()
        
        return True
    except Exception as e:
        print(f"Error sending to {recipient}: {str(e)}")
        return False

if __name__ == '__main__':
    app.run(debug=True)