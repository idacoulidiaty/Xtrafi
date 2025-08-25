import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

def send_reset_email(receiver_email, reset_link, sender_email, app_password):
    """
    Envoie un email pour réinitialisation de mot de passe
    """
    message = MIMEMultipart()
    message["From"] = sender_email
    message["To"] = receiver_email
    message["Subject"] = "Réinitialisation de votre mot de passe"

    body = f"Bonjour,\n\nCliquez sur ce lien pour réinitialiser votre mot de passe :\n{reset_link}\n\nSi vous n'avez pas demandé cette réinitialisation, ignorez ce message."
    message.attach(MIMEText(body, "plain"))

    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(sender_email, app_password)
            server.sendmail(sender_email, receiver_email, message.as_string())
        return True
    except Exception as e:
        print("Erreur envoi email :", e)
        return False

