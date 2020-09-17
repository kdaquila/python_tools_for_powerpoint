import mailer
import glob
import os


# TODO this still needs testing
def email_all_files(directory, from_addr, to_addr, subject, body, smtp_addr):
    pass
    # message = mailer.Message()
    #
    # message.From = from_addr
    #
    # message.To = to_addr
    #
    # message.Subject = subject
    #
    # message.Body = body
    #
    # files = glob.glob(os.path.join(directory, '*'))
    #
    # for f in files:
    #     print(f)
    #
    #     message.attach(f)
    #
    # sender = mailer.Mailer(smtp_addr)
    #
    # sender.send(message)
