require 'google_drive'
require 'mustache'
require 'redcarpet'

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

class EmailTemplate < Mustache
  DIRECTORY = File.join('.', 'posts')

  def _md_renderer
    @md_renderer ||= Redcarpet::Markdown.new(Redcarpet::Render::HTML.new)
  end

  def _file_contents(filename)
    File.open(File.join(DIRECTORY, filename), &:read)
  end

  def _to_html(extension, contents)
    if extension == '.md'
      puts "MD"
      _md_renderer.render(contents)
    else
      puts "HTML"
      contents
    end
  end

  def _read_from_file(filename)
    _to_html(File.extname(filename), _file_contents(filename))
  end

  def title
    _read_from_file('title.md')
  end
end
