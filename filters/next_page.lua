local first = true

function Header(el)
  if el.level == 1 then
    if first then
      first = false
      return nil
    end
    local next_page_section = [[
<w:p>
  <w:pPr>
    <w:sectPr>
      <w:type w:val="nextPage"/>
    </w:sectPr>
  </w:pPr>
</w:p>
]]
    return {pandoc.RawBlock('openxml', next_page_section), el}
  end
end
