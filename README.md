<<<<<<< HEAD
# JDE RPA

## Run

```powershell
pip install -r requirements.txt
playwright install chromium
python main.py
```

## Configuration

- `data/UserConfig.xlsx`: project name, base path, initialization processes and directory names.
- `data/Config.xlsx`: JDE URL, Windows generic credential name, environment, role and headless mode.
- Both use row 1 as headers and columns A/B as configuration key/value. Column C contains instructions.
- Relative base paths resolve from the framework root. Both configuration files are archived to the configuration directory.
- UserConfig values override matching Config values.
- Set jde_credential_name to an existing Windows Generic Credential target name. The process reads its username and password under the current Windows account.

## Flow

`main.py` -> `workflows/jde_flow/main_task.py` -> `web/JDE/jde_function.py`.
Browser lifecycle, task logging, failure screenshots and exit status are managed by the framework.
The login is followed by the six-module property-tax workflow. Pending modules are
recorded in `context.property_tax_status` (via `context.get`) and logged; they are
not completed downloads or journal entries. The framework task may finish normally
while the business workflow remains pending.

## Property tax workflow

1. Reuse the existing JDE login.
2. Ledger export awaits recording. For local testing, supply a ledger file.
3. Locate the template under the UserConfig input directory, populate `JDE明细账`
   by headers, and save `房产税计税表 公司 会计期间.xlsx` to the configured directory.
4. Update income totals by BU/document type from the audit worksheet and ledger;
   preserve the MOH row and update tax, total and journal-reference formulas.
5. Recalculate and compare the tax total against debit entries, and the income
   total directly against the negated Check total; reject mismatches.
6. Journal entry awaits field definitions and recorded JDE operations.

Add a `房产税配置` sheet to Config.xlsx, with headers in row 1 and keys/values
in columns A/B (UserConfig overrides matching values):

| Key | Value / behavior |
| --- | --- |
| 公司 | Required company identifier; use text to retain leading zeroes |
| 会计期间 | Required filename-safe period, e.g. 2026-10 |
| 输出目录 | Absolute path, or subdirectory relative to the framework output directory; blank uses that directory |
| 模板文件 | Optional filename relative to input directory, or absolute path; blank requires exactly one matching xlsx |
| 明细账文件 | Temporary local ledger path, relative to input directory or absolute |
| 明细账工作表 | Blank uses the first worksheet |
| 明细账标题行 | Default 1 |
| 模板标题行 | Default 1 in JDE明细账 |
| 标题映射 | Optional JSON object, e.g. `{"源科目": "科目"}`; blank matches identical headers |

Every nonblank source header must be mapped. Duplicate headers and destination
files are rejected. Template formulas are retained, but openpyxl does not calculate
them; Excel recalculation is requested on open. No comparison uses cached results.
The existing configuration workbooks have not been modified.

模块四按审核底稿的完整收入科目精确匹配（仅去除两端空格），
合并单元格 BU 读取左上角值，同一收入科目重复出现只计一次。
按收入科目、单据类型分别汇总，每条结果独占一行，即使 BU 相同也不合并；
计算底稿 B/C 列写 BU 和汇总金额乘以 -1。
按 BU、收入科目排列，同一科目的单据类型按降序排列。MOH 公式行保留在汇总行之后，
税额、合计、借项及贷项引用随行数更新；无匹配的旧金额引用归零。
Check 区 M:P 同步写入按完整科目及单据类型汇总的带符号金额。
当前支持示例布局（B9 起的汇总区、G:I 分录区及 M:P Check 区）。
模块五在核对前通过办公软件重算公式；分录决策仍待补充。

模块五通过 BU/借项金额/Check 金额表头和‘合计’标记动态定位，
支持明细及 Check 区行数变化。核对税额合计=借项汇总、
收入合计=Check合计乘以 -1（不扣除 250108），另检查 Check 合计等于其明细之和。
默认严格相等；`房产税配置.核对容差` 可显式配置非负金额，默认 0.00。
核对结果只写入日志与上下文（`property_tax_check_results`），不再生成 `.核对结果.json`。
任一项失败或公式结果缺失时停止，不进入分录；核对通过后仍等待分录决策规则。
`房产税配置.重算引擎` 默认为 `Excel.Application`，须安装支持 COM 的 Excel。
模块三先文件级复制模板，再由该引擎写入明细并重算保存，模板中的图表、
数据验证、条件格式等对象原样保留；配置为 `openpyxl` 时跳过 COM，但会丢失这些对象。
WPS 环境可配置 `ket.Application`，但需现场验证该版本支持独立实例及完整重算接口。
重算不更新外部链接；重算后的结果不可用时不会使用旧缓存放行。

BU 250108 固定为明细区最后一行，且仅出现一次，不从明细账另行汇总。
其 C 列保留 MOH LOOKUP 公式并固定引用会计期间 `$C$4`；模板缺失该行
或原值不是公式时，按示例生成 `=LOOKUP($C$4,MOH!$E$6:$AF$6,MOH!$E$7:$AF$7)`。
所有明细行（包括 250108）的 D 列生成 `=ROUND(C当前行*$C$8,2)`，
行数增减时重新生成合计范围及 250108 借项引用。

### 已记录的模板端标题

`config/property_tax_ledger_headers.json` 逐字记录输出示例
`data/房产税计税表 太古汇 202609.xlsx` 的 `JDE明细账!A1:AT1`，
共 46 列，数据从第 2 行开始。保留原始“帐”字、标题内空格及列序。
每列的 `template_header` 是模板端写法，`source_header` 暂为 null，
待实际导出的明细账提供后补充。模块三已读取此文件校验模板的 46 个标题，
并读取已填写的 `source_header` 作为源标题别名。
同名标题自动对应，`房产税配置.标题映射` 可单独配置异名列并覆盖默认别名，
方向为源标题到模板标题。按标题定位，支持源列顺序变化；标题内空格保留。
必须覆盖全部模板列，遇到缺列、未知源列、重复标题或多列写入同一列时终止。
写入前清空旧明细值，跳过全空行，保留文本前导零及首条明细的日期、金额等格式；
其他工作表及公式保留。实际源文件尚未提供，不推断异名列的业务对应关系。

## Framework modules

`core`, `config`, `log`, `init`, `web`, `client`, `data_processing`, and `utils` provide reusable framework components.
The original standalone project is kept separately for reference.

## Verification

```powershell
python -m unittest discover -s tests -v
```
=======
# tgh-jde-template



## Getting started

To make it easy for you to get started with GitLab, here's a list of recommended next steps.

Already a pro? Just edit this README.md and make it your own. Want to make it easy? [Use the template at the bottom](#editing-this-readme)!

## Add your files

- [ ] [Create](https://docs.gitlab.com/ee/user/project/repository/web_editor.html#create-a-file) or [upload](https://docs.gitlab.com/ee/user/project/repository/web_editor.html#upload-a-file) files
- [ ] [Add files using the command line](https://docs.gitlab.com/ee/gitlab-basics/add-file.html#add-a-file-using-the-command-line) or push an existing Git repository with the following command:

```
cd existing_repo
git remote add origin http://www.hanzhiai.com:7002/yky/tgh-jde-template.git
git branch -M main
git push -uf origin main
```

## Integrate with your tools

- [ ] [Set up project integrations](http://www.hanzhiai.com:7002/yky/tgh-jde-template/-/settings/integrations)

## Collaborate with your team

- [ ] [Invite team members and collaborators](https://docs.gitlab.com/ee/user/project/members/)
- [ ] [Create a new merge request](https://docs.gitlab.com/ee/user/project/merge_requests/creating_merge_requests.html)
- [ ] [Automatically close issues from merge requests](https://docs.gitlab.com/ee/user/project/issues/managing_issues.html#closing-issues-automatically)
- [ ] [Enable merge request approvals](https://docs.gitlab.com/ee/user/project/merge_requests/approvals/)
- [ ] [Set auto-merge](https://docs.gitlab.com/ee/user/project/merge_requests/merge_when_pipeline_succeeds.html)

## Test and Deploy

Use the built-in continuous integration in GitLab.

- [ ] [Get started with GitLab CI/CD](https://docs.gitlab.com/ee/ci/quick_start/index.html)
- [ ] [Analyze your code for known vulnerabilities with Static Application Security Testing (SAST)](https://docs.gitlab.com/ee/user/application_security/sast/)
- [ ] [Deploy to Kubernetes, Amazon EC2, or Amazon ECS using Auto Deploy](https://docs.gitlab.com/ee/topics/autodevops/requirements.html)
- [ ] [Use pull-based deployments for improved Kubernetes management](https://docs.gitlab.com/ee/user/clusters/agent/)
- [ ] [Set up protected environments](https://docs.gitlab.com/ee/ci/environments/protected_environments.html)

***

# Editing this README

When you're ready to make this README your own, just edit this file and use the handy template below (or feel free to structure it however you want - this is just a starting point!). Thanks to [makeareadme.com](https://www.makeareadme.com/) for this template.

## Suggestions for a good README

Every project is different, so consider which of these sections apply to yours. The sections used in the template are suggestions for most open source projects. Also keep in mind that while a README can be too long and detailed, too long is better than too short. If you think your README is too long, consider utilizing another form of documentation rather than cutting out information.

## Name
Choose a self-explaining name for your project.

## Description
Let people know what your project can do specifically. Provide context and add a link to any reference visitors might be unfamiliar with. A list of Features or a Background subsection can also be added here. If there are alternatives to your project, this is a good place to list differentiating factors.

## Badges
On some READMEs, you may see small images that convey metadata, such as whether or not all the tests are passing for the project. You can use Shields to add some to your README. Many services also have instructions for adding a badge.

## Visuals
Depending on what you are making, it can be a good idea to include screenshots or even a video (you'll frequently see GIFs rather than actual videos). Tools like ttygif can help, but check out Asciinema for a more sophisticated method.

## Installation
Within a particular ecosystem, there may be a common way of installing things, such as using Yarn, NuGet, or Homebrew. However, consider the possibility that whoever is reading your README is a novice and would like more guidance. Listing specific steps helps remove ambiguity and gets people to using your project as quickly as possible. If it only runs in a specific context like a particular programming language version or operating system or has dependencies that have to be installed manually, also add a Requirements subsection.

## Usage
Use examples liberally, and show the expected output if you can. It's helpful to have inline the smallest example of usage that you can demonstrate, while providing links to more sophisticated examples if they are too long to reasonably include in the README.

## Support
Tell people where they can go to for help. It can be any combination of an issue tracker, a chat room, an email address, etc.

## Roadmap
If you have ideas for releases in the future, it is a good idea to list them in the README.

## Contributing
State if you are open to contributions and what your requirements are for accepting them.

For people who want to make changes to your project, it's helpful to have some documentation on how to get started. Perhaps there is a script that they should run or some environment variables that they need to set. Make these steps explicit. These instructions could also be useful to your future self.

You can also document commands to lint the code or run tests. These steps help to ensure high code quality and reduce the likelihood that the changes inadvertently break something. Having instructions for running tests is especially helpful if it requires external setup, such as starting a Selenium server for testing in a browser.

## Authors and acknowledgment
Show your appreciation to those who have contributed to the project.

## License
For open source projects, say how it is licensed.

## Project status
If you have run out of energy or time for your project, put a note at the top of the README saying that development has slowed down or stopped completely. Someone may choose to fork your project or volunteer to step in as a maintainer or owner, allowing your project to keep going. You can also make an explicit request for maintainers.
>>>>>>> 61c26f814feec01d19dcd9f2caf41cd077163415
