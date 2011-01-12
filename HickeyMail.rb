# ====================================================================== #
# HickeyMail                                                             #
# An IMAP / E-Mail scanner to check for UConn School Closings and Delays #
# ver. 1.0                                                               #
# (C) 2010 - Rockwell B. Schrock & Erik J. Reynolds                      #
# THIS VERSION IS FOR MICROSOFT WINDOWS.                                 #
# ====================================================================== #

# You'll need to do:
# grab and install ruby from http://rubyforge.org/frs/download.php/69035/rubyinstaller-1.9.1-p378-rc2.exe
# run the following from an administrative command prompt: gem install win32-sapi
# you can then call the script via: ruby HickeyMail.rb
# or... set it as a task in Windows Task Scheduler (to run the above command)

# Require the IMAP libraries
require 'net/imap'
require 'time'
#require 'win32/sound'
require 'win32/sapi5'
include Win32

# Set our info
server = "imap.gmail.com"
port = 993
use_ssl = true
user = "root@dook.com"
pass = "HAHA"
folder = "INBOX"

# Connect to the GMail IMAP server
imap = Net::IMAP.new(server, port, use_ssl, nil, false)

# Login to GMail
imap.login(user, pass)

# Select the inbox
imap.select(folder)

# What is today?
todays_date = Time.now.localtime.strftime("%d-%b-%Y")

# Create our Text to Speech Engine
v = SpVoice.new

# store the value of teh message in a larger scope
hickey_message = ''

# Search for e-mails from Jay Hickey
imap.search(["NOT", "NEW", "SINCE", todays_date]).each do |msg_id|
  envelope = imap.fetch(msg_id, "ENVELOPE")[0].attr["ENVELOPE"]
   if envelope.from[0].mailbox == 'jay.hickey' and envelope.from[0].host = 'uconn.edu'
    
    (1..2).each do
      v.Speak("Attention.  Attention please.  Incoming message from Jay Hickey.")
    end
    
    fetch = imap.fetch(msg_id, 'BODY[TEXT]')[0]
    body = fetch.attr['BODY[TEXT]']
    paragraphs = body.split("\r\n\r\n")
    
    paragraphs.each do |paragraph|
      if paragraph.downcase.include?('due') and paragraph.downcase.include?('inclement')
        paragraph.gsub!("\r\n", '').gsub!('=', '')
        (1..2).each do
          v.Speak "#{paragraph}"
          hickey_message = paragraph
        end
        break
      end
    end
    #imap.store(message_id, "+FLAGS", [:Flagged])
  end
end

# Logout
puts "Logging out of GMail..."
imap.logout

# Disconnect from the server
puts "Disconnecting from the server..."
imap.disconnect
