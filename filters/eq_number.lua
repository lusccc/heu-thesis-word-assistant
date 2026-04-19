local chapter = 0
local eq_counter = 0
local eq_numbers = {}

local function ensure_chapter()
  if chapter == 0 then
    chapter = 1
  end
  return chapter
end

local function next_eq_number()
  local current = ensure_chapter()
  eq_counter = eq_counter + 1
  return string.format("（%d-%d）", current, eq_counter)
end

local function parse_eq_label(str)
  if not str then
    return nil
  end
  return str:match("{#(eq%-[^}]+)}")
end

local function is_display_math(elem)
  return elem.t == "Math" and elem.c and type(elem.c) == "table" and #elem.c >= 2 and elem.c[1] == "DisplayMath"
end

local function make_equation_block(math_elem, label)
  local number = next_eq_number()
  if label then
    eq_numbers[label] = number
  end
  
  -- 创建公式（编号由后处理脚本添加）
  -- 注意：Quarto 可能会在公式内部添加编号，但会被后处理脚本移除
  local math_copy = pandoc.Math(math_elem.c[1], math_elem.c[2])
  
  -- 创建段落内容：只包含公式
  local para_content = {math_copy}
  local para = pandoc.Para(para_content)
  
  if label then
    return pandoc.Div({para}, pandoc.Attr(label, {}, {["custom-style"]="equation-style"}))
  else
    return pandoc.Div({para}, pandoc.Attr("", {}, {["custom-style"]="equation-style"}))
  end
end

local function flush_buffer(buffer, blocks)
  if #buffer > 0 then
    table.insert(blocks, pandoc.Para(buffer))
    return {}
  end
  return buffer
end

function Header(el)
  if el.level == 1 then
    chapter = chapter + 1
    eq_counter = 0
  end
  return nil
end

function Para(el)
  local content = el.content
  local blocks = {}
  local buffer = {}
  local changed = false
  local i = 1
  while i <= #content do
    local node = content[i]
    if is_display_math(node) then
      buffer = flush_buffer(buffer, blocks)
      local label
      if i < #content and content[i + 1].t == "Str" then
        local candidate = parse_eq_label(content[i + 1].c)
        if candidate then
          label = candidate
          i = i + 1
        end
      end
      local eq_block = make_equation_block(node, label)
      table.insert(blocks, eq_block)
      changed = true
    else
      table.insert(buffer, node)
    end
    i = i + 1
  end
  buffer = flush_buffer(buffer, blocks)
  if changed then
    return blocks
  end
  return nil
end

function Cite(el)
  if #el.citations ~= 1 then
    return nil
  end
  local cite = el.citations[1]
  local id = cite.id
  if id and id:match("^eq%-") then
    local number = eq_numbers[id]
    if number then
      local target = "#" .. id
      return pandoc.Link({pandoc.Str(number)}, target, "", pandoc.Attr("", {"eq-ref"}))
    end
  end
  return nil
end
