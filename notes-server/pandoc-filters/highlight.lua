-- highlight.lua
-- This Pandoc Lua filter converts Word highlights and custom styles to HTML

-- Detect highlighted text
function Span(el)
    if el.attributes['highlight'] then
        -- convert highlight to <mark>
        return pandoc.RawInline('html', '<mark>' .. pandoc.utils.stringify(el) .. '</mark>')
    end
end

-- Convert bold/italic (if needed, Pandoc handles most by default)
function Strong(el)
    return el
end

function Emph(el)
    return el
end

-- Convert custom Citation style (example)
function Div(el)
    if el.attributes['style'] == 'Citation' then
        return pandoc.RawBlock('html', '<span class="citation">' .. pandoc.utils.stringify(el) .. '</span>')
    end
end