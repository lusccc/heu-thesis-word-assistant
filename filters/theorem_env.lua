-- theorem_env.lua
-- 处理数学环境：定义、定理、引理、推论、命题、例、注、假设、证明
-- 支持分章编号和交叉引用

-- 环境类型的中文名称
local env_names = {
  definition = "定义",
  theorem = "定理",
  lemma = "引理",
  corollary = "推论",
  proposition = "命题",
  example = "例",
  remark = "注",
  assumption = "假设"
}

-- 存储标签到编号的映射（用于交叉引用）
local label_map = {}

-- 处理整个文档（单一 pass）
function Pandoc(doc)
  local chapter = 0
  local counters = {
    definition = 0,
    theorem = 0,
    lemma = 0,
    corollary = 0,
    proposition = 0,
    example = 0,
    remark = 0,
    assumption = 0
  }
  
  -- 重置所有计数器
  local function reset_counters()
    for k, _ in pairs(counters) do
      counters[k] = 0
    end
  end
  
  -- 获取下一个编号
  local function next_number(env_type)
    if chapter == 0 then
      chapter = 1
    end
    counters[env_type] = counters[env_type] + 1
    return string.format("%d.%d", chapter, counters[env_type])
  end
  
  -- 处理定理类环境
  local function process_theorem_env(el, env_type, label)
    local number = next_number(env_type)
    local env_name = env_names[env_type]
    
    -- 存储标签映射
    if label and label ~= "" then
      label_map[label] = {
        number = number,
        name = env_name,
        full = env_name .. " " .. number
      }
    end
    
    -- 获取可选的环境名称（如"收敛定理"）
    local custom_name = el.attributes["name"]
    
    -- 构建标题
    local title_text
    if custom_name then
      title_text = string.format("%s %s（%s）", env_name, number, custom_name)
    else
      title_text = string.format("%s %s", env_name, number)
    end
    
    -- 创建标题（加粗）
    local title = pandoc.Strong({pandoc.Str(title_text)})
    
    -- 获取内容
    local content = el.content
    
    -- 构建新的块列表
    local blocks = {}
    
    -- 如果第一个块是段落，将标题插入到段落开头
    if #content > 0 and content[1].t == "Para" then
      local first_para = content[1]
      local new_content = {title, pandoc.Space()}
      for _, item in ipairs(first_para.content) do
        table.insert(new_content, item)
      end
      table.insert(blocks, pandoc.Para(new_content))
      
      -- 添加剩余的块
      for i = 2, #content do
        table.insert(blocks, content[i])
      end
    else
      -- 标题单独成段
      table.insert(blocks, pandoc.Para({title}))
      for _, block in ipairs(content) do
        table.insert(blocks, block)
      end
    end
    
    -- 返回带有标识符的 Div（用于交叉引用跳转）
    local attr = pandoc.Attr(label or "", {"theorem-env", env_type}, {})
    return pandoc.Div(blocks, attr)
  end
  
  -- 处理证明环境
  local function process_proof(el)
    local content = el.content
    local blocks = {}
    
    -- 创建"证明"标题（加粗）
    local title = pandoc.Strong({pandoc.Str("证明")})
    
    -- 如果第一个块是段落，将标题插入到段落开头
    if #content > 0 and content[1].t == "Para" then
      local first_para = content[1]
      local new_content = {title, pandoc.Space()}
      for _, item in ipairs(first_para.content) do
        table.insert(new_content, item)
      end
      table.insert(blocks, pandoc.Para(new_content))
      
      -- 添加剩余的块（除了最后一个）
      for i = 2, #content - 1 do
        table.insert(blocks, content[i])
      end
      
      -- 处理最后一个块，添加 QED 符号
      if #content > 1 then
        local last_block = content[#content]
        if last_block.t == "Para" then
          local last_content = {}
          for _, item in ipairs(last_block.content) do
            table.insert(last_content, item)
          end
          table.insert(last_content, pandoc.Space())
          table.insert(last_content, pandoc.Str("∎"))
          table.insert(blocks, pandoc.Para(last_content))
        else
          table.insert(blocks, last_block)
          table.insert(blocks, pandoc.Para({pandoc.Str("∎")}))
        end
      else
        -- 只有一个段落，在末尾添加 QED
        local last_block = blocks[#blocks]
        if last_block.t == "Para" then
          table.insert(last_block.content, pandoc.Space())
          table.insert(last_block.content, pandoc.Str("∎"))
        end
      end
    else
      -- 标题单独成段
      table.insert(blocks, pandoc.Para({title}))
      for i, block in ipairs(content) do
        if i == #content and block.t == "Para" then
          -- 最后一个段落，添加 QED
          local new_content = {}
          for _, item in ipairs(block.content) do
            table.insert(new_content, item)
          end
          table.insert(new_content, pandoc.Space())
          table.insert(new_content, pandoc.Str("∎"))
          table.insert(blocks, pandoc.Para(new_content))
        else
          table.insert(blocks, block)
        end
      end
      
      -- 如果没有内容，只添加 QED
      if #content == 0 then
        blocks[1].content[#blocks[1].content + 1] = pandoc.Space()
        blocks[1].content[#blocks[1].content + 1] = pandoc.Str("∎")
      end
    end
    
    return pandoc.Div(blocks, pandoc.Attr("", {"proof-env"}, {}))
  end
  
  -- 处理 Div 元素
  local function process_div(el)
    -- 检查是否是数学环境
    local env_type = nil
    for _, class in ipairs(el.classes) do
      if env_names[class] then
        env_type = class
        break
      elseif class == "proof" then
        env_type = "proof"
        break
      end
    end
    
    if not env_type then
      return nil
    end
    
    -- 获取标签（用于交叉引用）
    local label = el.identifier
    
    -- 处理证明环境（不编号）
    if env_type == "proof" then
      return process_proof(el)
    end
    
    -- 处理其他环境（需要编号）
    return process_theorem_env(el, env_type, label)
  end
  
  -- 第一遍：遍历文档块，处理 Div 并收集标签映射
  local new_blocks = {}
  for _, block in ipairs(doc.blocks) do
    if block.t == "Header" and block.level == 1 then
      chapter = chapter + 1
      reset_counters()
      table.insert(new_blocks, block)
    elseif block.t == "Div" then
      local result = process_div(block)
      if result then
        table.insert(new_blocks, result)
      else
        table.insert(new_blocks, block)
      end
    else
      table.insert(new_blocks, block)
    end
  end
  
  -- 第二遍：处理交叉引用
  local function process_cite(el)
    if #el.citations ~= 1 then
      return nil
    end
    
    local cite = el.citations[1]
    local id = cite.id
    
    -- 检查是否是定理环境的引用
    local info = label_map[id]
    if info then
      local target = "#" .. id
      return pandoc.Link({pandoc.Str(info.full)}, target, "", pandoc.Attr("", {"theorem-ref"}))
    end
    
    return nil
  end
  
  -- 遍历所有块，处理交叉引用
  local final_blocks = {}
  for _, block in ipairs(new_blocks) do
    local new_block = pandoc.walk_block(block, {Cite = process_cite})
    table.insert(final_blocks, new_block)
  end
  
  return pandoc.Pandoc(final_blocks, doc.meta)
end
