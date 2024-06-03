require 'google_drive'

EMAIL_INDEX = 2
SPREADSHEET_ID = "1Yr-5JKmD5-q9G427oJ4_HUFsid3Mmm2m8Gt9pMFcTxk"

def emails(worksheet)
  (2..worksheet.num_rows).map { |row|
    worksheet[row, EMAIL_INDEX]
  }.reject { |e|
    e.nil? || e.strip.empty?
  }
end

def subscribed_emails
  session = GoogleDrive::Session.from_service_account_key('fiddle-newsletter-0cd0c5d4f02c.json')
  sub_ws, unsub_ws = session.spreadsheet_by_key(SPREADSHEET_ID).worksheets
  emails(sub_ws) - emails(unsub_ws)
end
