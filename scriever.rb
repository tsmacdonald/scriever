require 'dotenv/load'
require 'google_drive'
require "mailersend-ruby"

class GoogleEmailList
  def initialize(google_keyfile, email_index, spreadsheet_id)
    @google_key = google_keyfile
    @email_index = email_index.to_i
    @spreadsheet_id = spreadsheet_id
  end

  def _emails(worksheet)
    (2..worksheet.num_rows).map { |row|
      worksheet[row, @email_index]
    }.reject { |e|
      e.nil? || e.strip.empty?
    }
  end

  def subscribed_emails
    session = GoogleDrive::Session.from_service_account_key(@google_key)
    sub_ws, unsub_ws = session.spreadsheet_by_key(@spreadsheet_id).worksheets
    _emails(sub_ws) - _emails(unsub_ws)
  end
end

class Newsletter
  def initialize(template_id)
    google_key = ENV['SCRIEVER_GOOGLE_KEY'] || File.join('.', 'google-service-account.json')
    email_index = ENV['SCRIEVER_EMAIL_INDEX']
    spreadsheet_id = ENV['SCRIEVER_SPREADSHEET']
    @email_list = GoogleEmailList.new(google_key, email_index, spreadsheet_id)

    @template_id = template_id
    @mailersend_token = ENV['SCRIEVER_MAILERSEND_TOKEN']
    @from_address = ENV['SCRIEVER_FROM_ADDRESS']
    @from_name = ENV['SCRIEVER_FROM_NAME']
  end

  def ms_client
    @ms_client ||= Mailersend::Client.new(@mailersend_token)
  end

  def send!(dry_run)
    emails = dry_run ? [ENV['SCRIEVER_TEST_EMAIL']] : @email_list.subscribed_emails

    puts "Preparing to send an email to #{emails.count} emails, starting with #{emails.first}"
    puts 'Continue? (yes/no)'
    if STDIN.gets.chomp.downcase != 'yes'
      puts 'Aborting'
      return
    end

    puts 'Subject:'
    subject = STDIN.gets.chomp

    ms_bulk_email = Mailersend::BulkEmail.new(ms_client)
    ms_bulk_email.messages = emails.map do |to_email|
      {
        'from' => { 'email' => @from_address,
                    'name' => @from_name },
        'to' => [{ 'email' => to_email }],
        'subject' => subject,
        'template_id' => @template_id,
        'personalization' => [
          {
            'email' => to_email,
            'data' => {
              'account_name' => 'Tim Macdonald'
            }
          }
        ]
      }
    end
    response = ms_bulk_email.send
    puts response
    r_id = response['bulk_email_id']
    puts "Sent! (#{r_id})"
    status(r_id)
  end

  def status(bulk_email_id)
    Mailersend::BulkEmail.new(ms_client).get_bulk_status(bulk_email_id:).body.tap do |b|
      puts b
    end
  end
end

if __FILE__ == $PROGRAM_NAME
  template_id = ARGV.first || (raise 'Please specify a template_id')
  dry_run = (ARGV[1] != 'yes-i-really-want-to-send-an-email')

  Newsletter.new(template_id).send!(dry_run)
end
