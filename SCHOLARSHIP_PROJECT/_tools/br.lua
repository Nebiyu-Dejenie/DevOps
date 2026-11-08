-- Converts raw HTML <br> (used for line breaks inside Markdown table cells) into real line breaks in Word output.
function RawInline(el)
  if el.format:match("html") and el.text:match("^<br%s*/?>$") then
    return pandoc.LineBreak()
  end
end
