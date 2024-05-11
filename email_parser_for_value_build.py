from datetime import datetime, timedelta
from email.message import EmailMessage
import email
import imaplib
import pandas as pd
import re
import ssl
import smtplib


current_datetime = datetime.now()
current_date = current_datetime.date()
current_year = current_datetime.year
current_month = current_datetime.month
current_day = current_datetime.day
date_object = datetime(current_year, current_month, current_day)
current_week_of_year = date_object.isocalendar()[1]
next_year = current_year + 1


blank1 = []
blank2 = []
blank3 = []
blank4 = []
blank5 = []
blank6 = []
guest_date_received = []
guest_name = []
guest_email_address = []
guest_check_in = []
guest_alt_check_in = []
guest_number_of_nights = []
guest_six_and_over = []
guest_under_six = []
guest_which_house = []
guest_contact_wish = []
guest_message = []


IMAP_SERVER = '***********'
EMAIL = '******************.com' # the inbox we look in
PASSWORD = '*********.'  # inbox pw
SENDER_EMAIL = '****************.com' # from this account


date_12_hours_ago = (datetime.now() - timedelta(hours=12)).strftime('%d-%b-%Y')

mail = imaplib.IMAP4_SSL(IMAP_SERVER)
mail.login(EMAIL, PASSWORD)
mail.select('inbox')


def get_guest_full_name(guest_email):
    match = re.search(r'Full Name\s*:\s*(.*?)\n.*?\nEmail', guest_email, re.DOTALL)
    clean = match.group(0)
    cleaned = str(clean).replace('Full Name : ', '').replace('Email', '').rstrip('\n').lstrip()
    if cleaned:
        return cleaned
    else:
        return ' '


def japanese_get_guest_full_name(guest_email):
    match = re.search(r'氏名（平仮名でご入力ください）:\s*(.*?)\n', guest_email, re.DOTALL)
    clean = match.group(0)
    cleaned = str(clean).replace('氏名（平仮名でご入力ください）: ', '')
    if cleaned:
        return cleaned
    else:
        return ' '


def get_guest_email(guest_email):
    match = re.search(r'Email\s*:\s*(.*?)\n.*?\nCheck-in', guest_email, re.DOTALL)
    clean = match.group(0)
    cleaned = str(clean).replace('Email : ', '').replace('Check-in', '').rstrip('\n').lstrip()
    if cleaned:
        return cleaned
    else:
        return ' '


def japanese_get_guest_email(guest_email):
    match = re.search(r'メールアドレス:\s*(.*?)\n', guest_email, re.DOTALL)
    clean = match.group(0)
    cleaned = str(clean).replace('メールアドレス: ', '')
    if cleaned:
        return cleaned
    else:
        return ' '


def get_guest_checkin(guest_email):
    match = re.search(r'Check-in Date:\s*(.*?)\n', guest_email, re.DOTALL)
    clean = match.group(0)
    cleaned = str(clean).replace('Check-in Date: ', '').replace('Is', '').rstrip('\n').lstrip()
    if cleaned:
        return cleaned
    else:
        return ' '


def japanese_get_guest_checkin(guest_email):
    match = re.search(r'チェックイン希望日をご記入ください \(dd/mm/yyyy\) ::\s*(.*?)\n', guest_email, re.DOTALL)
    clean = match.group(0)
    cleaned = str(clean).replace('チェックイン希望日をご記入ください (dd/mm/yyyy) :: ', '')
    if cleaned:
        return cleaned
    else:
        return ' '


def get_guest_alt_checkin(guest_email):
    match = re.search(r'Is\s*(.*?)\n.*?\nNumber', guest_email, re.DOTALL)
    clean = match.group(0)
    cleaned = str(clean).replace('Is your check-in date flexible?: ', '').replace('Number', '').rstrip('\n').lstrip()
    if cleaned:
        return cleaned
    else:
        return ' '


def japanese_get_guest_alt_checkin(guest_email):
    match = re.search(r'上記希望日以外でのご旅行が可能な場合はその他チェックイン可能日をご記入ください \(dd/mm/yyyy\) ::\s*(.*?)\n', guest_email, re.DOTALL)
    clean = match.group(0)
    cleaned = str(clean).replace('上記希望日以外でのご旅行が可能な場合はその他チェックイン可能日をご記入ください (dd/mm/yyyy) :: ', '')
    if cleaned:
        return cleaned
    else:
        return ' '


def get_guest_number_of_nights(guest_email):
    match = re.search(r'Number\s*(.*?)\n.*?\nHow many guests', guest_email, re.DOTALL)
    clean = match.group(0)
    cleaned = str(clean).replace('Number of nights: ', '').replace('How many guests', '').rstrip('\n').strip()
    if cleaned:
        return cleaned
    else:
        return ' '


def japanese_get_guest_number_of_nights(guest_email):
    match = re.search(r'ご希望の最低宿泊日数を記入してください ::\s*(.*?)\n', guest_email, re.DOTALL)
    clean = match.group(0)
    cleaned = str(clean).replace('ご希望の最低宿泊日数を記入してください :: ', '')
    if cleaned:
        return cleaned
    else:
        return ' '


def get_guest_six_and_over(guest_email):
    match = re.search(r'How\smany\sguests\s*(.*?)\n', guest_email, re.DOTALL)
    clean = match.group(0)
    cleaned = clean.replace('How many guests (6 years old and over) in your group? : ', '')
    if cleaned:
        return cleaned
    else:
        return ' '


def japanese_get_guest_six_and_over(guest_email):
    match = re.search(r'宿泊人数\(小学生以上） ::s*(.*?)\n', guest_email, re.DOTALL)
    clean = match.group(0)
    cleaned = str(clean).replace('宿泊人数(小学生以上） :: ', '')
    if cleaned:
        return cleaned
    else:
        return ' '


def get_guest_under_six(guest_email):
    match = re.search(r'How many children \(0-5 years old\) will share a bed with an adult in your group\?\s*:\s*(.*?)\n', guest_email, re.DOTALL)
    clean = match.group(0)
    cleaned = clean.replace('How many children (0-5 years old) will share a bed with an adult in your group?  : ', '')
    if cleaned:
        return cleaned
    else:
        return ' '


def japanese_get_guest_under_six(guest_email):
    match = re.search(r'未就学児で添い寝のお子様の数 （※3名目からは『宿泊人数（小学生以上）』に追加ください::\s*(.*?)\n', guest_email, re.DOTALL)
    clean = match.group(0)
    cleaned = str(clean).split(':')[2]
    if cleaned:
        return cleaned
    else:
        return ' '


def get_guest_which_house(guest_email):
    match = re.search(r'Which\s*(.*?)\n.*?\nDo', guest_email, re.DOTALL)
    clean = match.group(0)
    cleaned = clean.replace('Which machiya house would you like to book? : ', '').replace('Do', '')
    if cleaned:
        return cleaned
    else:
        return ' '


def japanese_get_guest_which_house(guest_email):
    match = re.search(r'ご希望の施設やお部屋はありますか？:\s*(.*?)\n', guest_email, re.DOTALL)
    clean = match.group(0)
    cleaned = str(clean).replace('ご希望の施設やお部屋はありますか？: ', '')
    if cleaned:
        return cleaned
    else:
        return ' '


def get_guest_contact_wish(guest_email):
    match = re.search(r'days\s*(.*?)\n.*?\nMessage:', guest_email, re.DOTALL)
    clean = match.group(0)
    cleaned = str(clean).split(":", 1)[1].replace('Message:', '').strip()
    if cleaned:
        return cleaned
    else:
        return ' '


def japanese_get_guest_contact_wish(guest_email):
    match = re.search(r'ご希望の宿泊日の一部の日付が利用可能になった場合、ご連絡を受けたいですか？:\s*(.*?)\n', guest_email, re.DOTALL)
    clean = match.group(0)
    cleaned = str(clean).split(":", 1)[1]
    if cleaned:
        return cleaned
    else:
        return ' '


def get_guest_message(guest_email):
    match = re.search(r'Message:\s*([\s\S]*?)(?:\n\s*\n|$)', guest_email)
    if match:
        cleaned = match.group(1).strip()
        return cleaned
    else:
        return ''


def main():
    result, data = mail.search(None, '(FROM "{}" SINCE "{}")'.format(SENDER_EMAIL, date_12_hours_ago))
    if result == 'OK':
        for num in data[0].split():
            result, data = mail.fetch(num, '(RFC822)')
            if result == 'OK':
                raw_email = data[0][1]
                msg = email.message_from_bytes(raw_email)

                for part in msg.walk():
                    if part.get_content_type() == 'text/plain':
                        guest_email = part.get_payload(decode=True).decode('utf-8')
                        for line in guest_email.splitlines():
                            if line.startswith("Full"):
                                guestname = get_guest_full_name(guest_email)
                                guest_name.append(guestname)
                            if line.startswith("氏名"):
                                guestname = japanese_get_guest_full_name(guest_email)
                                guest_name.append(guestname)
                            if line.startswith("Email"):
                                guestemailaddress = get_guest_email(guest_email)
                                guest_email_address.append(guestemailaddress)
                            if line.startswith("メールアドレス:"):
                                guestemailaddress = japanese_get_guest_email(guest_email)
                                guest_email_address.append(guestemailaddress)
                            if line.startswith("Check-in"):
                                guestcheckin = get_guest_checkin(guest_email)
                                guest_check_in.append(guestcheckin)
                            if line.startswith("チェックイン希望日をご記入ください"):
                                guestcheckin = japanese_get_guest_checkin(guest_email)
                                guest_check_in.append(guestcheckin)
                            if line.startswith("Is"):
                                guestaltcheckin = get_guest_alt_checkin(guest_email)
                                guest_alt_check_in.append(guestaltcheckin)
                            if line.startswith("上記希望日以外でのご旅行が可能な場合はその他チェックイン可能日をご記入ください"):
                                guestaltcheckin = japanese_get_guest_alt_checkin(guest_email)
                                guest_alt_check_in.append(guestaltcheckin)
                            if line.startswith("Number"):
                                guestnumberofnights = get_guest_number_of_nights(guest_email)
                                guest_number_of_nights.append(guestnumberofnights)
                            if line.startswith("ご希望の最低宿泊日数を記入してください"):
                                guestnumberofnights = japanese_get_guest_number_of_nights(guest_email)
                                guest_number_of_nights.append(guestnumberofnights)
                            if line.startswith("How many guests"):
                                guestsixandover = get_guest_six_and_over(guest_email)
                                guest_six_and_over.append(guestsixandover)
                            if line.startswith("宿泊人数(小学生以上）"):
                                guestsixandover = japanese_get_guest_six_and_over(guest_email)
                                guest_six_and_over.append(guestsixandover)
                            if line.startswith("How many children"):
                                guestundersix = get_guest_under_six(guest_email)
                                guest_under_six.append(guestundersix)
                            if line.startswith("未就学児で添い寝のお子様の数"):
                                guestundersix = japanese_get_guest_under_six(guest_email)
                                guest_under_six.append(guestundersix)
                            if line.startswith("Which"):
                                guestwhichhouse = get_guest_which_house(guest_email)
                                guest_which_house.append(guestwhichhouse)
                            if line.startswith("ご希望の施設やお部屋はありますか？:"):
                                guestwhichhouse = japanese_get_guest_which_house(guest_email)
                                guest_which_house.append(guestwhichhouse)
                            if line.startswith("Do you"):
                                guestcontactwish = get_guest_contact_wish(guest_email)
                                guest_contact_wish.append(guestcontactwish)
                            if line.startswith("ご希望の宿泊日の一部の日付が利用可能になった場合、ご連絡を受けたいですか？:"):
                                guestcontactwish = japanese_get_guest_contact_wish(guest_email)
                                guest_contact_wish.append(guestcontactwish)
                            if line.startswith("Message:"):
                                guestmessage = get_guest_message(guest_email)
                                guest_message.append(guestmessage)

                        blank1.append(' ')
                        blank2.append(' ')
                        blank3.append(' ')
                        blank4.append(' ')
                        blank5.append(' ')
                        blank6.append(' ')
                        date_received_str = msg['Date']
                        date_received = datetime.strptime(date_received_str, '%a, %d %b %Y %H:%M:%S %z')
                        guest_date_received.append(date_received.strftime('%Y-%m-%d %H:%M:%S'))
                        if len(guest_name) > len(guest_message):
                            if not guest_email[0].startswith("氏名") or not guest_email[0].startswith("Full"):
                                guestmessage = guest_email
                                guest_message.append(guestmessage)
                            else:
                                guestmessage = " "
                                guest_message.append(guestmessage)

                mail.store(num, '+FLAGS', '/Seen')

    else:
        print("Failed to fetch emails.")

    data = {
        "blank1": blank1,
        "blank2": blank2,
        "blank3": blank3,
        "blank4": blank4,
        "Date Received": guest_date_received,
        "Name": guest_name,
        "Email": guest_email_address,
        "Check-in Date": guest_check_in,
        "Flexible Check-in": guest_alt_check_in,
        "Number of Nights": guest_number_of_nights,
        "Guests (6 and over)": guest_six_and_over,
        "Children (under 6)": guest_under_six,
        "House Preference": guest_which_house,
        "Contact Preference": guest_contact_wish,
        "blank5": blank5,
        "blank6": blank6,
        "Message": guest_message
    }

    df = pd.DataFrame(data)

    try:
        print(guest_email)
    except UnboundLocalError:
        print('No guest emails')

    df.to_excel(f'latest emails.xlsx', index=False)
    mail.close()
    mail.logout()

    # If you wish to send the table of data as an Excel file you can customise the section below,
    # otherwise you can just remove this block

    email_sender = '********************.com'
    email_password = '***************'
    email_receivers = ['***************.com', '****************.com']
    email_bcc = ['*************.com']
    subject = "Kyoto Machiya Waiting List Emails"
    body = f"""
    <html>
        <body>
            <img src="https://kyotomachiyas.com/wp-content/uploads/2022/10/Logo-Transparent-Logo.png" width="25%" height="25%">
            <br>
            <h1>Kyoto Machiya Waiting List Emails</h1>
            <h2>From the past 12 hours</h2>
        </body>
    </html>
    """
    data_file = f"latest emails.xlsx"
    em = EmailMessage()
    em['From'] = email_sender
    em['To'] = email_receivers
    em['Subject'] = subject
    em.set_content(body, subtype='html')
    with open(data_file, "rb") as f:
        em.add_attachment(f.read(), filename=data_file, maintype="application", subtype="octet-stream")
    context = ssl.create_default_context()
    with smtplib.SMTP_SSL('smtp.gmail.com', 465, context=context) as smtp:
        smtp.login(email_sender, email_password)
        smtp.sendmail(email_sender, email_receivers + email_bcc, em.as_string())


if __name__ == '__main__':
    main()
