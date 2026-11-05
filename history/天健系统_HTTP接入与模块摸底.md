# 天健系统 · HTTP 接入与模块摸底（经验边记）

性质：边做边写的接入笔记。不是 Skill，不替代正式授权；凭据只放本机，不入库。  
对照入口：[天健门户](https://ifly.csair.com/#/portal/index)  
过程产物：`_tmp/probe_tianjian_*.py`、`_tmp/tianjian_modules_probe.json`、`_tmp/tianjian_all_menu_pc_outline.json` 等。

---

## 一、网页怎么进

1. 业务站：`https://ifly.csair.com/`（SPA，标题「天健系统」）  
2. 门户路由：`https://ifly.csair.com/#/portal/index`  
3. 未登录会跳南航 IAM（南航 e 家）：  
   `https://iam.csair.com/idp/oauth2/authorize?client_id=iFLY&redirect_uri=https%3A%2F%2Fifly.csair.com%2F%23%2Flogin%3Ftarget%3D%252Fportal%252Findex&response_type=code&state=...`  
4. IAM 登录成功后带 OAuth `code` 回天健 `/#/login`，前端调 `doLoginIam` 换票，写入 Cookie `IFLY_TOKEN`。  
5. 关联：`IEB` 主机在壳配置里为 `https://ieb.csair.com`；天健与 IEB／飞行门户不是同一套 Cookie。

---

## 二、HTTP 鉴权参数（已踩坑确认）

### 2.1 必带

| 项 | 放哪 | 说明 |
| --- | --- | --- |
| Cookie | 请求 Cookie 头 | 至少含 `IFLY_TOKEN=...`；从 cURL 的 `-b` 整段拷最稳 |
| **IFLY-TOKEN** | **请求头（横杠）** | 值与 Cookie 里 JWT 相同。**只塞 Cookie 或写成 `IFLY_TOKEN` 头会 401** |
| RESOURCE-ID | 请求头 | 常见 `bnVsbA==`（`null` 的 base64）；按菜单权限点进具体页时可能换成资源节点 id |
| Referer / Origin | 请求头 | 建议 `https://ifly.csair.com/` |
| Accept | 请求头 | `application/json, text/plain, */*` |

### 2.2 怎么拿到有效凭据

1. 浏览器先登录进门户（确认不是 IAM 登录页）。  
2. F12 → Network → 找 `https://ifly.csair.com/api/login-app/loginUserInfo`（或任意已 200 的 `/api/...`）。  
3. 右键 **Copy as cURL**，整段贴进仓库根目录 `天健cookies.txt`（可多行）。  
4. **不要**只靠 Console `document.cookie`：往往缺头、或 HttpOnly／请求头 `IFLY-TOKEN` 拷不全。  
5. `csn2api` 对 IEB／OA／SOC 的「cURL 粘贴验证」方法可照搬，但仓内**没有**天健／ifly 现成适配；天健要用本仓 `_tmp` 探针。

### 2.3 JWT 里有什么

`IFLY_TOKEN` 为 HS256 JWT，载荷可见字段示例：`username`（员工号）、`usernameChn`、`loginType`、`isOauth`、`ts`（签发时间）。  
服务端会校验；过期或未带 `IFLY-TOKEN` 头 → 业务接口 401。

### 2.4 本机文件约定

- 凭据：`天健cookies.txt`（gitignore 习惯：勿提交）  
- 探针：`_tmp/probe_tianjian_portal.py`、`_tmp/probe_tianjian_modules.py`、`_tmp/probe_tianjian_auth_variants.py`

---

## 三、前端壳与后端应用（架构一眼）

壳是 Vue SPA + 微前端 WebComponent，版本号 `window.webComponentVersion`。  
子包路径示例：

| 模块 code | 名称 | 静态路径 |
| --- | --- | --- |
| iflyCommon | 公共 | `/ifly-common/iflyCommonWebComponent` |
| iflyPortal / iflyAdmin | 天健门户／管理 | `/ifly-web/...` |
| fleet | 机队 | `/third-party/fleet/...` |
| train | 翔翼 | `/third-party/train/...` |
| irms | 安全 | `/third-party/irms/...` |
| plm | 麦图 | `/third-party/plm/...` |

业务 API 同域前缀：`https://ifly.csair.com/api/{xxx-app}/...`  
已见后端应用根：`login-app`、`manage-app`、`qual-app`、`train-manage-app`、`school-app`、`label-app`、`profile-app`、`info-manage-app`、`os-app`、`irms-app`、`fleet-manage-app`、`fleet-duty-app`、`gateway-app` 等。

---

## 四、菜单两套数：全量 vs 我可见

| 接口 | 含义 | 本轮结果（账号 295494） |
| --- | --- | --- |
| `GET /api/login-app/loginUserInfo/frontResouce` | **当前账号有权**菜单树（注意拼写 Resou**ce**） | 约 **69** 节点 |
| `GET /api/login-app/loginUserInfo/frontResourceAll` | 系统资源全量（扁平，需按 `parentId` 组树） | 约 **4217** 节点 |
| `GET /api/login-app/favorite/commonUseResourceList` | 常用收藏 | 本轮仅「飞行任务查询」 |

全量 PC 端一级（挂在「天健系统PC浏览器端」下，与《天健功能菜单2026》同族）：

1. 天健首页  
2. 我的工作  
3. 人力资源规划  
4. 学员培养  
5. 训练管理（旧）  
6. 训练管理  
7. 资质管理  
8. 运行管理  
9. 机队管理  
10. 安全管理  
11. 飞行员管理  
12. 退出管理  
13. 公共管理  
14. 后台管理  

明细 L2：`_tmp/tianjian_all_menu_pc_outline.json`。

本账号（货航飞行部训练管理室／资质管理）**实际可见**主要是：天健首页若干卡片、训练管理（旧／新）局部、运行管理→飞行任务查询、安全管理局部、飞行员管理→健康消息、公共管理→权限申请／系统指引、机队运行→绩效等——**不是**全量十四个一级都开。  
排查「系统有没有某模块」看 `frontResourceAll`；排查「我能不能调」看 `frontResouce` + 接口 403／1010。

---

## 五、用户信息接口

`GET /api/login-app/loginUserInfo`

返回 `data` 里常用字段：`staffNum`、`staffName`、`orgunitName`、`position`、`gender`、各类 base／fleet 标签、`systemManager`／`permManager` 等。  
本轮样例：员工号 295494，张峻哲，中国南方航空货运有限公司飞行部训练管理室，岗位「资质管理」。

---

## 六、按模块摸接口（边试边记）

说明：下列「参数」以本轮试探为准；很多列表要 **GET + query**，误用 POST 会返回业务码 400「Request method 'GET' is not supported」（HTTP 仍可能是 200，看 JSON 的 `code`）。  
无权限常见：`code=403`「您没有权限」，或 `code=1010`「无权限操作此记录」。

### 6.1 login-app（登录／首页／菜单）

| 方法 | 路径 | 参数／注意 | 本轮 |
| --- | --- | --- | --- |
| GET | `/api/login-app/loginUserInfo` | 可加 `?r=时间戳` | 200／业务 200 |
| GET | `/api/login-app/loginUserInfo/frontResouce` | 有权菜单 | 200 |
| GET | `/api/login-app/loginUserInfo/frontResourceAll` | 全量资源 | 200，4217 条 |
| GET | `/api/login-app/favorite/commonUseResourceList` | 常用 | 200 |
| GET | `/api/login-app/homepage/user/config/getTheme` | 主题 | 200 |
| GET | `/api/login-app/homepage/user/config/getConfig` | 首页配置 | 200 |
| GET | `/api/login-app/report/loginCountStatic` | 登录统计 | 200 |
| POST | `/api/login-app/login/doLoginIam` | body 需 IAM 回调 **code**；缺 code 报「code不能为空」 | 通、业务失败 |

### 6.2 manage-app（待办／权限申请／基础）

| 方法 | 路径 | 参数／注意 | 本轮 |
| --- | --- | --- | --- |
| GET | `/api/manage-app/manage/work/page` | query：`pageNum`、`pageSize`（**不要 POST**） | 200，list 可空 |
| GET | `/api/manage-app/manage/work/approvedCount` | 已办计数 | 200 |
| GET | `/api/manage-app/perm/apply/page` | `pageNum`、`pageSize` | 200 |
| GET | `/api/manage-app/perm/apply/options` | 申请选项 | 200 |
| GET | `/api/manage-app/perm/apply/permManage/list/leader` | 要「申请类型」等，缺则 1001 | 参数不全 |

网页：`/portal/mw/todoWork/...`、`/portal/permissionManage/permissionApply`。

### 6.3 os-app（运行／飞行任务）

| 方法 | 路径 | 参数／注意 | 本轮 |
| --- | --- | --- | --- |
| GET | `/api/os-app/flightTask/options` | 机场等选项字典 | 200 |
| GET | `/api/os-app/flightTask/page` | 仅日期不够；还要**航班号，或出发+到达机场** | 1010 提示缺航班／机场 |
| GET | `/api/os-app/flightTask/crewInfo` | `staffNum`；还受数据权限约束 | 本号查自己 1010「没有查询该员工号权限」 |
| GET | `/api/os-app/task` | 要 `type`、开始／结束日期、`staffNum` 等 | 缺 type→1001 |

网页：`/portal/os/taskQueryManage/taskQuery`。  
经验：任务类接口参数比「有菜单」更严，且**数据范围权限**和菜单权限是两层。

### 6.4 train-manage-app（训练）

本账号对多条 `train-manage-app/...` 试探为 **403**。  
菜单上可见「训练安排／训练过程／转升路径」等路由，不代表列表 API 已授权。  
继续摸时：从门户包里按路由名（如 `trainPreArrangement`）搜邻近 `/train-manage-app/...`，用有权账号的 cURL 对照。

### 6.5 qual-app（资质）

本轮抽样 `autorenew`、`aptitude` 等多为 **403**。  
全量菜单有「资质管理」一级；货航资质岗账号仍可能未开具体 API——以后用「有权账号」的 Network 补参数表。

### 6.6 school-app（学员／航校）

全量菜单有完整「学员培养」；本账号 API 多为 403／1010。  
与货航是否建设学员模块的产品决策分开：系统侧能力在全量资源里已在。

### 6.7 label-app（作风量化／帮扶标签等）

量化报表、帮扶在库等抽样 **403** 或要求 POST。  
桌面功能菜单里的作风量化，本号未开 API。

### 6.8 profile-app（一人一本／画像）

缺员工号或无权限时：`1001`／`1010`。  
导出类接口勿瞎 GET。

### 6.9 info-manage-app／irms-app／fleet-*／gateway-app

- 系统指引文档列表：`GET .../operation/document/guideList`、`.../page` 本轮 200、list 空。  
- 体检／FSOP 等：403 或要 POST。  
- `GET /api/gateway-app/hi-agent/agents`：200，返回 AI 助手 id 映射。  
- 机队 `seat/queryUser` 等：本轮 403。

---

## 七、业务响应码习惯（经验）

| code | 大致意思 |
| --- | --- |
| 200 | 业务成功（仍要看 `data`） |
| 400 | 常为方法不对（GET／POST 用错） |
| 403 | 无此接口／菜单权限 |
| 1001 | 参数校验失败（缺字段） |
| 1010 | 无数据权限或不允许操作该记录 |
| 500 | 服务端未知错误／缺参导致 |

HTTP 状态与业务 `code` 常不一致：以 JSON 为准。

---

## 八、和货航建设方案的关系（备忘，非拍板）

1. 产品上已是「一壳多模块多 *-app」；全量菜单覆盖学员、资质、训练、运行等，货航可裁剪开通。  
2. 隔离诉求优先落在**数据／账号／权限域**（及是否分库），不必先等同于「再写一套天健前端」。  
3. 本号可见菜单很少，说明**权限与组织**已能大幅减负；分库是更强数据隔离，裁菜单是产品裁剪。

---

## 九、续写清单（下次接着做）

1. 用「训练／资质真有权」账号各导一条 cURL，补 `train-manage-app`、`qual-app` 列表的真实 query／body。  
2. 飞行任务：补齐「航班号或 OD 机场 + 日期」最小成功样例，写入本节 6.3。  
3. 把 `frontResourceAll` 组树后的 L2／L3 与桌面《天健功能菜单2026》黄标表逐项对照，产出「系统有／账号无／货航可能不需要」三列。  
4. 探针脚本合并：解析 cURL → 验票 → 拉菜单 → 按模块试探 → 更新本稿。

---

## 十、本轮命令备忘

```text
python _tmp/probe_tianjian_modules.py
```

产出主索引：`_tmp/tianjian_modules_probe.json`。  
凭据失效时重新 Copy as cURL 覆盖 `天健cookies.txt` 再跑。


---

## 十一、数据库／数据域怎么组织的（从前端与账号能看见的）

说明：本轮**拿不到**服务端 DDL／真正的库表清单（无库账号、管理端多数接口 403）。下面结论来自：前端包、用户信息、权限文案、公司编码，属于「逻辑结构推断」，不是 DBA 拍板。

### 11.1 已经能确定的逻辑模型

1. **一套产品、多公司编码**，不是「货航从零另起域名」那种空白系统。  
   - 管理端公司枚举：CZ＝南航，CS＝南货航，OQ＝重航。  
   - 训练／人事基地枚举里：CSG＝南货航。  
   - 当前登录用户：orgunitBase=CSG，orgunitBaseChn=货航，部门全称含「中国南方航空货运有限公司…」。  
2. **隔离主要靠权限，不是靠前端里的 tenantId／分库开关**。  
   - 前端全文几乎搜不到 	enantId、multiTenant、分库、sharding。  
   - 明确有「数据权限」「基地权限」「航校数据权限」「角色权限」；越权常见业务码 403（菜单／接口）与 1010（无权限操作此记录／查某员工）。  
3. **后端按 *-app 拆服务**（login／qual／train-manage／os／school…），表很大概率按域分散在多服务库或同集群多 schema——这是微服务常态；**不等于**已经按航司物理分库。  
4. UI 里的 dataSource 多指「系统生成 vs 手动新增」这类业务来源，**不是**动态数据源／分库配置。  
5. 文案已承认货航数据路径差异：例如作风量化「南航为天健，货航、重航为飞行门户」——说明历史上货航有一部分仍走门户，天健侧用公司维度区分，而不是物理上不存在货航。

### 11.2 三种改造路径对照（客观）

| 方案 | 含义 | 和现状匹配度 | 代价／风险 |
| --- | --- | --- | --- |
| A. 同库加强权限（逻辑隔离） | 继续共库，用 CSG／角色／数据权限把货航看住 | 与现网模型一致；已有公司码 | 隔离强度最弱；误配权限会串数；审计／故障域仍共用 |
| B. **同产品分库（或独立实例）** | 代码／模块同宗，货航独立库或独立部署，菜单裁剪 | 匹配「多 app 同壳」；满足「数据完全隔离」且不必重写业务 | 要迁数、切连接、管双环境发布；但升级仍可跟主线 |
| C. 镜像一套「货航版」（库＋应用双轨） | 再复制一摊库表＋可能改分支代码，长期两套演进 | 像项目组说的「完全隔离的货航版」 | 双倍回归、口径漂移、规章一变改两处；最贵 |
| D. 重开新系统／新库模型 | 抛开天健表结构重做 | 与已存在的 CS／CSG 编码、十四个一级模块重复建设 | 最慢、最贵；除非监管强制「不得复用」 |

「镜像一个」若只指 **库结构复制一份给货航实例用（B）**，合理；若指 **镜像后当永久分叉产品（C）**，不合理（除非有强合规）。  
「重开一个新的（D）」：从前端已看到的模块完整度与公司编码看，**不适合作为默认**。  
「分库（B）」：在要「完全隔离」时，是和现架构最合拍的硬隔离手段。

### 11.3 综合判断（给拍板用的一句话）

现网天健是 **逻辑多公司（CZ／CS／CSG／OQ）＋RBAC／数据权限** 的大一统产品；货航不是外人，是库里的一个 CSG／CS 维度。  
若只要责任清晰、功能裁剪：优先 **同宗＋权限／菜单裁剪（A）**，必要时再上 **独立库／实例（B）**。  
**不推荐**默认走「重开新库模型」或「镜像后永久双轨货航版」；后者只在合规明确要求物理隔离且不愿同版本跟随时再上。

### 11.4 仍缺的硬证据（下次向项目组要）

1. 各 *-app 的实际库名／是否已分 schema。  
2. 有无按 hrBase／company 的行级拦截中间件说明。  
3. 货航若分库：主数据（人员、组织）是否与南航共享只读库还是全量拷贝。  
4. 与 IEB／排班／翔翼的库表或接口边界图。


（完 · 续写随时追加章节）
