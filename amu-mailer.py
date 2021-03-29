import argparse
import getpass
import itertools
import os.path
import smtplib
import yaml

from email.mime.text import MIMEText

ARGPARSER = argparse.ArgumentParser()
ARGPARSER.add_argument('-d', '--dry-run', dest='dry_run', action='store_true',
                       help='Do not actually send e-mails.')
ARGPARSER.add_argument('-f', '--from', dest='from_', action='store',
                       help='Sender\'s e-mail adress.')
ARGPARSER.add_argument('-m', '--messages', dest='messages', action='store', default='messages.yaml',
                       help='YAML file with configuration and messages.')
ARGPARSER.add_argument('-v', '--verbose', dest='verbose', action='store_true')
ARGS = ARGPARSER.parse_args()

SMTP_HOST = "smtp.univ-amu.fr"
SMTP_PORT = 587

# Stats
total, sent = 0, 0

try:
    # Confirm input file exists.
    if not os.path.exists(ARGS.messages):
        raise('File "' + ARGS.messages + '" not found.')

    # Load messages file.
    messages = list(yaml.safe_load_all(open(ARGS.messages, "r")))

    configuration = messages[0]

    # Confirm sender and subject are specified.
    from_ = ARGS.from_ or configuration['from']
    subject = configuration['subject']
    if from_ is None or subject is None:
        raise('Fields "from" and "subject" are required. Specify them as program arguments or in the first document of "message')

    # Update SMTP paramter from default if specified.
    SMTP_HOST = SMTP_HOST or configuration['smtp_host']
    SMTP_PORT = SMTP_PORT or configuration['smtp_port']

    # Prompt for user if not specified.
    user = configuration['user'] or getpass.getuser()

    # Prompt for password.
    mot_de_passe = getpass.getpass()

    server = None
    total = len(messages) - 1

    # Iterate over all messages.
    for message in messages[1:]:

        try:
            # (Re-)connect to SMTP server if needed.
            if server is None:
                    # Instantiate server.
                server = smtplib.SMTP(host=SMTP_HOST, port=SMTP_PORT)
                if ARGS.verbose:
                    print('(Re-)connected to STMP server ' + SMTP_HOST)

                # Login to server.
                server.starttls()
                server.login(user, mot_de_passe)
                if ARGS.verbose:
                    print('Logged in to SMTP server ' + SMTP_HOST)

            # Craft message.
            mime = MIMEText(configuration['header'] + '\n' + message['body'] + '\n' + configuration['footer'], 'plain', 'utf-8')
            mime['Subject'] = subject
            mime['From'] = from_
            mime['To'] = message['to'] if isinstance(message['to'], str) else ','.join(message['to'])
            if 'cc' in configuration :
                mime['Cc'] = message['cc'] if isinstance(message['cc'], str) else ','.join(message['cc'])
            if 'bcc' in configuration :
                mime['Bcc'] = message['bcc'] if isinstance(message['bcc'], str) else ','.join(message['bcc'])

            if ARGS.verbose:
                print('New message:')
                # mime.as_string()
                # print(mime.as_string().encode())

            # Send message.
            if not ARGS.dry_run:
                echecs = server.send_message(mime)
                if echecs:
                    try:
                        import termcolor
                        termcolor.cprint('FAILED: ' + echecs[0], 'red', attrs=['blink'])
                    except:
                        print('FAILED: ' + echecs[0])
                else:
                    sent += 1
                    if ARGS.verbose:
                        print('Message successfully sent.')

        except smtplib.SMTPAuthenticationError as e:
            raise e
        except smtplib.SMTPResponseException as e:
            if ARGS.verbose:
                print("SMTP operation failed: [" + str(e.smtp_code) + "] " + str(e.smtp_error))
            server.quit()
            server = None

    if ARGS.verbose:
        print(str(sent) + '/' + str(total) + ' e-mails sent')

except smtplib.SMTPAuthenticationError as e:
    print("Authentication failed: [" + str(e.smtp_code) + "] " + str(e.smtp_error))
except RuntimeError as e:
    print(e.args)
