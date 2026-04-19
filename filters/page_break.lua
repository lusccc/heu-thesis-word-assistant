-- 使用 nextPage 分节符：每章从新页开始（不一定是奇数页）
local first_chapter = true

function Header(el)
  if el.level == 1 then
    if first_chapter then
      first_chapter = false
      -- 第一章不插入分节符，因为目录后已经有分节符了
      return el
    else
      -- 从第二章开始，在标题前插入分节符（下一页）
      local section_break = pandoc.RawBlock('openxml', [[
<w:p>
  <w:pPr>
    <w:sectPr>
      <w:type w:val="nextPage"/>
    </w:sectPr>
  </w:pPr>
</w:p>
]])
      return {section_break, el}
    end
  end
  return nil
end
