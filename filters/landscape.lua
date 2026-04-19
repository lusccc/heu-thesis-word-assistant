-- landscape.lua
-- 处理带有 .landscape class 的 Div，将其内容放在横向（landscape）页面中
-- 用法：在 thesis.qmd 中给需要横向的div添加 .landscape class
--   ::: {#tbl-xxx .landscape}
--   表格内容
--   :::

-- A4页面参数（与文档纵向设置一致）
local PAGE_W = "11906"     -- A4宽度（缇）
local PAGE_H = "16838"     -- A4高度（缇）
local MARGIN_TOP = "1587"
local MARGIN_BOTTOM = "1587"
local MARGIN_LEFT = "1417"
local MARGIN_RIGHT = "1417"
local MARGIN_HEADER = "1134"
local MARGIN_FOOTER = "1134"

-- 纵向分节符（结束前一个纵向节）
local function portrait_section_break()
  return string.format([[
<w:p>
  <w:pPr>
    <w:sectPr>
      <w:type w:val="nextPage"/>
      <w:pgSz w:w="%s" w:h="%s"/>
      <w:pgMar w:top="%s" w:bottom="%s" w:left="%s" w:right="%s" w:header="%s" w:footer="%s" w:gutter="0"/>
    </w:sectPr>
  </w:pPr>
</w:p>
]], PAGE_W, PAGE_H, MARGIN_TOP, MARGIN_BOTTOM, MARGIN_LEFT, MARGIN_RIGHT, MARGIN_HEADER, MARGIN_FOOTER)
end

-- 横向分节符（结束横向节）
local function landscape_section_break()
  return string.format([[
<w:p>
  <w:pPr>
    <w:sectPr>
      <w:type w:val="nextPage"/>
      <w:pgSz w:w="%s" w:h="%s" w:orient="landscape"/>
      <w:pgMar w:top="%s" w:bottom="%s" w:left="%s" w:right="%s" w:header="%s" w:footer="%s" w:gutter="0"/>
    </w:sectPr>
  </w:pPr>
</w:p>
]], PAGE_H, PAGE_W, MARGIN_TOP, MARGIN_BOTTOM, MARGIN_LEFT, MARGIN_RIGHT, MARGIN_HEADER, MARGIN_FOOTER)
end

function Div(el)
  if el.classes:includes('landscape') then
    local before = pandoc.RawBlock('openxml', portrait_section_break())
    local after = pandoc.RawBlock('openxml', landscape_section_break())

    local blocks = pandoc.List({before})
    blocks:extend(el.content)
    blocks:insert(after)
    return blocks
  end
end
