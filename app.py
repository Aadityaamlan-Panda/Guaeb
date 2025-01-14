import os
import json
import tkinter as tk
from tkinter import filedialog, messagebox, Text
from azure.ai.contentsafety import ContentSafetyClient
from azure.ai.contentsafety.models import AnalyzeImageOptions, ImageData, AnalyzeTextOptions
from azure.ai.textanalytics import TextAnalyticsClient
from azure.core.credentials import AzureKeyCredential
from sendgrid import SendGridAPIClient
from sendgrid.helpers.mail import Mail
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import smtplib
import cohere

KEY = "YOUR_COGNITIVE_SERVICES_KEY"  # Replace with your actual key
ENDPOINT = "YOUR_COGNITIVE_SERVICES_ENDPOINT"  # Replace with your actual endpoint
LANGUAGE_KEY = "YOUR_TEXT_ANALYTICS_KEY"  # Replace with your actual key
LANGUAGE_ENDPOINT = "YOUR_TEXT_ANALYTICS_ENDPOINT"  # Replace with your actual endpoint
CONFIG_FILE = "email_config.json"

COHERE_API_KEY = "YOUR_COHERE_API_KEY"  # Replace with your actual API key
cohere_client = cohere.ClientV2(COHERE_API_KEY)

# Utility functions to save and load email configuration
def save_config(data):
    with open(CONFIG_FILE, "w") as file:
        json.dump(data, file)

def load_config():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r") as file:
            return json.load(file)
    return {}

# Azure Content Safety Analysis Functions
def analyze_text_moderation(client, text):
    try:
        response = client.analyze_text(AnalyzeTextOptions(text=text))
        restricted_categories = [
            f"{category.category} (Severity: {category.severity})"
            for category in response.categories_analysis if category.severity > 0
        ]
        return restricted_categories
    except Exception as e:
        return [f"Error analyzing text: {e}"]

def analyze_image_moderation(client, image_path):
    try:
        with open(image_path, "rb") as file:
            request = AnalyzeImageOptions(image=ImageData(content=file.read()))
            response = client.analyze_image(request)
            restricted_categories = [
                f"{category.category} (Severity: {category.severity})"
                for category in response.categories_analysis if category.severity > 0
            ]
            return restricted_categories
    except Exception as e:
        return [f"Error analyzing image: {e}"]

# Azure Text Analytics (Sentiment Analysis)
def analyze_sentiment(client, text):
    try:
        documents = [text]
        result = client.analyze_sentiment(documents, show_opinion_mining=True)
        output = []
        for document in result:
            if document.is_error:
                return [f"Error analyzing sentiment: {document.error}"]
            output.append(f"Overall Sentiment: {document.sentiment}")
            output.append(
                f"Positive={document.confidence_scores.positive:.2f}, "
                f"Neutral={document.confidence_scores.neutral:.2f}, "
                f"Negative={document.confidence_scores.negative:.2f}"
            )
            for sentence in document.sentences:
                output.append(
                    f"Sentence: {sentence.text}\n"
                    f"  Sentiment: {sentence.sentiment}\n"
                    f"  Positive={sentence.confidence_scores.positive:.2f}, "
                    f"Neutral={sentence.confidence_scores.neutral:.2f}, "
                    f"Negative={sentence.confidence_scores.negative:.2f}"
                )
        return output
    except Exception as e:
        return [f"Error analyzing sentiment: {e}"]

def analyze_formality(email_text):
     try:
        # Use the Cohere API Chat endpoint for formality detection
    
        response = cohere_client.chat(
            model="command-r-plus-08-2024",
            messages = [{"role": "user", "content": f"Determine if the following text is formal or informal:\n\n'{email_text}'"}]
        )
        # Extract the formality result from the response
        result = response.message.content[0].text
        return [f"Formality Analysis: {result}"]
     except Exception as e:
        return [f"Error analyzing formality: {e}"]

# SMTP Mailtrap Configuration
MAILTRAP_SMTP_SERVER = "live.smtp.mailtrap.io"
MAILTRAP_SMTP_PORT = 587
MAILTRAP_USERNAME = "api"
MAILTRAP_PASSWORD = "YOUR_MAILTRAP_PASSWORD"

def send_email(sender_email, receiver_email, subject, body, attachment_path):
    try:
        # Create a multipart message and set headers
        message = MIMEMultipart()
        message["From"] = sender_email
        message["To"] = receiver_email
        message["Subject"] = subject

        # Attach body to email
        message.attach(MIMEText(body, "plain"))

        # If attachment exists, attach it
        if attachment_path:
            with open(attachment_path, "rb") as attachment:
                file_name = os.path.basename(attachment_path)
                part = MIMEBase("application", "octet-stream")
                part.set_payload(attachment.read())
                encoders.encode_base64(part)
                part.add_header("Content-Disposition", f"attachment; filename={file_name}")
                message.attach(part)

        # Send email using Mailtrap's SMTP server
        with smtplib.SMTP(MAILTRAP_SMTP_SERVER, MAILTRAP_SMTP_PORT) as server:
            server.starttls()
            server.login(MAILTRAP_USERNAME, MAILTRAP_PASSWORD)
            server.sendmail(sender_email, receiver_email, message.as_string())

        messagebox.showinfo("Success", "Email sent successfully!")
    except Exception as e:
        print(e)
        messagebox.showerror("Error", f"Failed to send email: {e}")

# GUI Implementation
class EmailInterface:
    def __init__(self, root):
        self.root = root
        self.root.title("GUAEB EMAIL BOT")
        self.root.geometry("800x600")

        # Load configuration
        self.config = load_config()

        # Azure Clients
        self.content_safety_client = ContentSafetyClient(ENDPOINT, AzureKeyCredential(KEY))
        self.text_analytics_client = TextAnalyticsClient(
            endpoint=LANGUAGE_ENDPOINT, credential=AzureKeyCredential(LANGUAGE_KEY)
        )

        # Text Area
        tk.Label(root, text="Compose Email", font=("Arial", 14)).pack(pady=10)
        self.email_text = Text(root, wrap=tk.WORD, font=("Arial", 12))
        self.email_text.pack(padx=10, pady=5, fill=tk.BOTH, expand=True)

        # Image Upload
        tk.Label(root, text="Attachments:", font=("Arial", 12)).pack(pady=5)
        self.image_path_var = tk.StringVar()
        tk.Entry(root, textvariable=self.image_path_var, font=("Arial", 12), width=50).pack(padx=10, pady=5, side=tk.LEFT)
        tk.Button(root, text="Browse", command=self.browse_image).pack(padx=5, side=tk.LEFT)

        # Analyze Button
        tk.Button(root, text="Analyze Content", command=self.analyze_content, bg="blue", fg="white", font=("Arial", 12)).pack(pady=10, padx=5, side=tk.LEFT)

        # Email Sending
        tk.Button(root, text="Send Email", command=self.send_email_interface, bg="green", fg="white", font=("Arial", 12)).pack(pady=10, padx=5, side=tk.LEFT)


    def browse_image(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("All Files", "*.*"), ("Image Files", "*.jpg;*.png;*.jpeg")]
        )
        self.image_path_var.set(file_path)


     
    def analyze_content(self):
        email_text = self.email_text.get("1.0", tk.END).strip()
        file_path = self.image_path_var.get()

        sentiment_results = []
        text_issues = []
        image_issues = []
        formality_results = []

        if email_text:
            sentiment_results.extend(analyze_sentiment(self.text_analytics_client, email_text))
            text_issues.extend(analyze_text_moderation(self.content_safety_client, email_text))
            formality_results.extend(analyze_formality(email_text))

        if file_path and file_path.lower().endswith(('.jpg', '.png', '.jpeg')):
            image_issues.extend(analyze_image_moderation(self.content_safety_client, file_path))

        output = "\n\n".join([
            "Sentiment Analysis Results:\n" + "\n".join(sentiment_results) if sentiment_results else "No sentiment analysis performed.",
            "Formality Analysis Results:\n" + "\n".join(formality_results) if formality_results else "No formality analysis performed.",
            "Text Moderation Issues:\n" + "\n".join(text_issues) if text_issues else "No text moderation issues detected.",
            "Image Moderation Issues:\n" + "\n".join(image_issues) if image_issues else "No image moderation issues detected.",
        ])

        self.display_output_popup(output)

    def display_output_popup(self, output):
        # Create a new popup window
        popup = tk.Toplevel(self.root)
        popup.title("Analysis Results")
        popup.geometry("600x400")

        # Add a text widget with a scrollbar for displaying the output
        output_text = Text(popup, wrap=tk.WORD, font=("Arial", 12))
        output_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar = tk.Scrollbar(popup, orient=tk.VERTICAL, command=output_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        output_text.config(yscrollcommand=scrollbar.set)
        output_text.insert(tk.END, output)

        # Disable editing in the output text widget
        output_text.config(state=tk.DISABLED)

        # Add a close button to the popup
        tk.Button(popup, text="Close", command=popup.destroy, bg="red", fg="white").pack(pady=10)

    def send_email_interface(self):
        email_window = tk.Toplevel(self.root)
        email_window.title("Send Email")
        email_window.geometry("400x400")

        tk.Label(email_window, text="Sender Email:").pack(pady=5)
        sender_email_entry = tk.Entry(email_window)
        sender_email_entry.insert(0, "sender@demomailtrap.com")
        sender_email_entry.pack(pady=5)

        tk.Label(email_window, text="Recipient Email:").pack(pady=5)
        receiver_email_entry = tk.Entry(email_window)
        receiver_email_entry.pack(pady=5)

        tk.Label(email_window, text="Subject:").pack(pady=5)
        subject_entry = tk.Entry(email_window)
        subject_entry.pack(pady=5)

        def send():
            sender_email = sender_email_entry.get()
            receiver_email = receiver_email_entry.get()
            subject = subject_entry.get()
            body = self.email_text.get("1.0", tk.END).strip()
            attachment = self.image_path_var.get()

            send_email(sender_email, receiver_email, subject, body, attachment)

        tk.Button(email_window, text="Send", command=send, bg="green", fg="white").pack(pady=10)

# Run Application
if __name__ == "__main__":
    root = tk.Tk()
    app = EmailInterface(root)
    root.mainloop()
