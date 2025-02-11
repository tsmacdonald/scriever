require 'dotenv/load'
require 'google_drive'
require "mailersend-ruby"
require 'mustache'
require 'redcarpet'

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

class EmailTemplate < Mustache
  def initialize(template)
    super
    self.template_file = File.join(__dir__, 'email-templates', template)
  end

  def render_post(post_dir)
    @post_dir = post_dir
    render
  end

  def _md_renderer
    @md_renderer ||= Redcarpet::Markdown.new(Redcarpet::Render::HTML.new)
  end

  def _file_contents(filename)
    raise 'Post dir not set' unless @post_dir

    File.open(File.join(@post_dir, filename), &:read)
  end

  def _to_html(extension, contents)
    if extension == '.md'
      _md_renderer.render(contents.strip)
    else
      contents
    end
  end

  def _read_from_file(filename)
    _to_html(File.extname(filename), _file_contents(filename))
  end

  def title
    _read_from_file('title.txt')
  end

  def header
    _read_from_file('header.md')
  end

  def body
    _read_from_file('body.md')
  end

  def footer
    _read_from_file('footer.md')
  end
end

class Newsletter
  def initialize(post_dir, template = nil)
    @post_dir = post_dir
    @google_key = ENV['SCRIEVER_GOOGLE_KEY'] || File.join('.', 'google-service-account.json')
    @template = template || ENV['SCRIEVER_TEMPLATE']
    @spreadsheet_id = ENV['SCRIEVER_SPREADSHEET']
    @email_index = ENV['SCRIEVER_EMAIL_INDEX']
    @mailersend_token = ENV['SCRIEVER_MAILERSEND_TOKEN']
    @from_address = ENV['SCRIEVER_FROM_ADDRESS']
    @from_name = ENV['SCRIEVER_FROM_NAME']
  end

  def ms_client
    @ms_client ||= Mailersend::Client.new(@mailersend_token)
  end

  def send!(dry_run = true)
    # emails = GoogleEmailList.new(@google_key, @email_index, @spreadsheet_id).subscribed_emails
    emails = ['mailersend@tsmacdonald.com']
    msg = EmailTemplate.new(@template).render_post(@post_dir)
    if dry_run
      spacer = '-' * 80
      puts "Sending the following email:\n#{spacer}\n#{msg}\n#{spacer}\nto:\n#{emails.join(',')}"
    else
      ms_bulk_email = Mailersend::BulkEmail.new(ms_client)

      ms_bulk_email.messages = emails.map do |to_email|
        puts "Preparing an email to #{to_email}"
        {
          'from' => { 'email' => @from_address, 'name' => @from_name },
          'to' => [{ 'email' => to_email }],
          'subject' => 'Hello world!',
          # TODO: 'text' =>
          'template_id' => 'k68zxl2qvy3lj905',
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
      puts 'SENDING AN EMAIL!!!'
      puts ms_bulk_email.messages
      response = ms_bulk_email.send
      puts "Response: #{response}"
      puts 'status:'
      puts status(response['bulk_email_id']).body
    end
  end

  def status(bulk_email_id)
    Mailersend::BulkEmail.new(ms_client).get_bulk_status(bulk_email_id:)
  end
end

if __FILE__ == $PROGRAM_NAME
  post_dir = ARGV.first || (raise 'Please specify a post to send')
  template = ARGV[1]
  dry_run = (ARGV[2] != 'yes-i-really-want-to-send-an-email')

  Newsletter.new(post_dir, template).send!(dry_run)
end
