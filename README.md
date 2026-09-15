# dpsk

个人工作区仓库。只收录项目代码与配置，排查过程产生的临时文件（驱动 / UAC / 代理探测脚本、解包产物、浏览器 profile）一律不入库，详见 [`.gitignore`](.gitignore)。

## 目录结构

| 路径 | 说明 |
| --- | --- |
| `study-site/` | 英语学习工作台（考研二 + 四六级网页版）。**独立仓库，以 git submodule 形式引入** |
| `static-host/` | 静态托管页面与部署指南 |
| `.dsh/skills/` | DSH harness 技能包（`grill-me`、`impeccable` 等） |
| `study-progress.json` | 学习进度数据 |

## 克隆

`study-site` 是子模块，需要递归拉取：

```bash
git clone --recursive https://github.com/QINGyi71771/dpsk.git
```

已经克隆过、但子模块目录为空时：

```bash
git submodule update --init --recursive
```

## 子模块说明

`study-site/` 拥有自己的仓库 [QINGyi71771/study-site](https://github.com/QINGyi71771/study-site)，在本仓库中只记录一个 commit 指针。修改学习站代码请：

```bash
cd study-site
git add -A && git commit -m "..."
git push
```

回到本仓库后，再提交更新后的指针：

```bash
cd ..
git add study-site
git commit -m "chore: 更新 study-site 指针"
```

## 本机 Git 环境备注

本机的 Windows 原生 TLS（schannel）存在故障，报 `schannel: AcquireCredentialsHandle failed: SEC_E_NO_CREDENTIALS`，导致 `curl` / `Invoke-WebRequest` 无法建立 HTTPS 连接（`gh` 因自带 Go TLS 栈不受影响）。

因此本仓库的 `.git/config` 中已设置：

```
http.sslBackend = openssl
```

使用 PortableGit（`D:\dpsk\.tools\git`）时如遇同样的 TLS 报错，需在全局配置中同样指定 `http.sslBackend=openssl`。
