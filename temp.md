我要在当前目录里创建一个新项目，EstimateM，用于生成"项目配置项规模及工作量估算记录"Excel文件。                                       
                                                                                                                                
  ## 项目背景                                                                                                     
  这是一套工具集的第二个系统，第一个系统是 ConfigM（配置管理记录生成），放在同级目录 ../ConfigM/ 下。         
  ## 架构决策（已确定，不再讨论）             
                                                                                                                                                
  每个系统完全独立： 
  - EstimateM 维护自己的 SQLite 数据库 estimatem.db 
  - 用户向 EstimateM 导入 Excel 文件，EstimateM 只解析自己需要的 Sheet 和列
  - 不读取 ConfigM 的数据库，不依赖 ConfigM 服务
  - Excel 文件是数据的权威来源，各系统从同一份完整 Excel 各取所需 
  ## 技术栈（已确定，不再讨论）        
  与 ConfigM 完全一致：                                                                                              
  - Python + FastAPI + SQLite（aiosqlite）+ Alpine.js + openpyxl                
  - 启动方式：python3 run.py，使用系统 Python，不使用 venv 
  - 前端：单页 Alpine.js 应用，风格与 ConfigM 一致   

  ## ConfigM 现有的 Excel 
  Excel 文件目前包含 4 个 Sheet：                                                                                                               
  1. 项目信息：项目名称、项目代号、项目组长、配置管理员                                    
  2. 项目进度计划：WBS、任务名称、计划开始时间、计划完成时间、关联配置项    
  3. 配置项清单：序号、配置项名称、配置项标识、版本
  4. 基线组成：基线名称、基线标识、基线版本、包含配置项名称、配置项标识、配置项版本                                                                                                                                      
  EstimateM 需要复用其中的部分数据，并在 Excel 中新增估算相关的列或 Sheet。                                      
  ## CLAUDE.md 要求                                                              
  请参考以下结构创建 CLAUDE.md，和 ConfigM 项目保持同一风格：
  1. 项目上下文（每次新会话必读）：项目定位、技术栈、关键文档路径 
  2. 当前阶段（已完成 / 下一步）                                           
  3. 关键文件清单 
  4. 已知规则和约定                           
  5. 会话开始时的操作                                                   
  6. 里程碑更新规则
  7. Skill routing，内容如下，直接复制写入 CLAUDE.md：                                                                                          
                                                                                                                                                
  ## Skill routing                                                                                                                              
                                                                                                                                                
  When the user's request matches an available skill, invoke it via the Skill tool. When in doubt, invoke the skill.                            
                                          
  Key routing rules:                                                                                                                            
  - Product ideas/brainstorming → invoke /office-hours                                               
  - Strategy/scope → invoke /plan-ceo-review
  - Architecture → invoke /plan-eng-review                                                                 
  - Design system/plan review → invoke /design-consultation or /plan-design-review
  - Full review pipeline → invoke /autoplan                                                        
  - Bugs/errors → invoke /investigate                                    
  - QA/testing site behavior → invoke /qa or /qa-only
  - Code review/diff check → invoke /review
  - Visual polish → invoke /design-review                                                  
  - Ship/deploy/PR → invoke /ship or /land-and-deploy                                          
  - Save progress → invoke /context-save                                                         
  - Resume context → invoke /context-restore       
  ## 现在请执行       
  先不要写任何代码，也不要创建任何文件。     
  请触发 /office-hours，围绕以下问题展开需求澄清：                                                                 
  - EstimateM 需要生成的估算记录包含哪些内容，有哪些维度的数据
  - 哪些数据来自 Excel 导入（ConfigM 已有字段），哪些是估算专有字段需要用户在 UI 中录入
  - 页面需要展示哪些信息，以什么形式呈现
  - 导出的 Excel 有哪些 Sheet，每个 Sheet 的列名和数据来源
  - 用户的完整操作流程（导入 → 录入 → 查看 → 导出）                                        
  - 导出文件的命名规则和存放位置