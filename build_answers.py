# -*- coding: utf-8 -*-
import docx, json, re

TEMPLATE = "/Users/xxx/code/test1/答案模板.docx"
OUT_DOCX = "/Users/xxx/code/test1/答案.docx"
OUT_JSON = "/Users/xxx/code/test1/答案.json"
PKG = "pkg-08e82c8d"
PLAYER = "陈焕炜"
CANNOT = "该需求不能直接转换为FOFA搜索语句"

# ---- FOFA query per 题号 ----
A = {}

# M001 指定IP
A["M001-S009"] = 'ip="20.247.40.92"'
# M002 C段网段
A["M002-S004"] = 'cidr="124.255.254.0/24"'
# M003 主域
A["M003-S002"] = 'domain="zendesk.com"'
# M004 网站名(host)
A["M004-S003"] = 'host="www.amazon.com"'
# M005 域名包含vpn
A["M005-S004"] = 'domain="*vpn*"'
# M006 英国
A["M006-S002"] = 'country="GB"'
# M007 韩国江原道
A["M007-S007"] = 'country="KR" && region="Gangwon"'
# M008 德国巴伐利亚州Munich
A["M008-S009"] = 'country="DE" && region="Bavaria" && city="Munich"'
# M009 ASN
A["M009-S008"] = 'asn="4134"'
# M010 荷兰 + 亚马逊云
A["M010-S009"] = 'country="NL" && cloud_name="Amazon"'
# M011 河北 + 华为云
A["M011-S005"] = 'region="Hebei" && cloud_name="HuaWeiCloud"'
# M012 开放3000端口
A["M012-S009"] = 'port="3000"'
# M013 标题含Jenkins或正文jenkins,非中国,2026-08-01之后
A["M013-S003"] = '(title="Jenkins" || body="jenkins") && country!="CN" && after="2026-08-01"'
# M014 中国DNS服务
A["M014-S002"] = 'country="CN" && protocol="dns"'
# M015 问句非检索 -> 不可转换
A["M015-S001"] = CANNOT
# M016 SSH协议且端口22022
A["M016-S006"] = 'protocol="ssh" && port="22022"'
# M017 TCP协议
A["M017-S001"] = 'base_protocol="tcp"'
# M018 SNMP服务不限端口
A["M018-S006"] = 'protocol="snmp"'
# M019 Memcached, 美国
A["M019-S010"] = 'protocol="memcached" && country="US"'
# M020 真正跑在TCP上的FTP
A["M020-S005"] = 'protocol="ftp" && base_protocol="tcp"'
# M021 真正SSH且banner含OpenSSH_10.4
A["M021-S010"] = 'protocol="ssh" && banner="OpenSSH_10.4"'
# M022 日本海康威视摄像头
A["M022-S007"] = 'country="JP" && app="海康威视-视频监控"'
# M023 恰好两级单字符子域 -> FOFA不能精确表达单字符通配 -> 不可转换
A["M023-S002"] = CANNOT
# M024 产品版本10.0.0不限产品
A["M024-S005"] = 'version="10.0.0"'
# M025 中间件(连接应用与底层基础服务的中间层软件)
A["M025-S010"] = 'category="Middleware"'
# M026 正文含NoSuchUpload
A["M026-S002"] = 'body="NoSuchUpload"'
# M027 响应含Content-Security-Policy-Report-Only
A["M027-S008"] = 'header="Content-Security-Policy-Report-Only"'
# M028 状态码308
A["M028-S011"] = 'status_code="308"'
# M029 响应含SSH-2.0-OpenSSH_6.1
A["M029-S010"] = 'banner="SSH-2.0-OpenSSH_6.1"'
# M030 header含Set-Cookie:PHPSESSID
A["M030-S003"] = 'header="Set-Cookie: PHPSESSID"'
# M031 引用jquery.min.js
A["M031-S009"] = 'body="jquery.min.js"'
# M032 荷兰阿姆斯特丹554端口且响应含200 OK
A["M032-S010"] = 'country="NL" && city="Amsterdam" && port="554" && banner="200 OK"'
# M033 GlobalSign签发且证书未过期
A["M033-S003"] = 'cert.issuer="GlobalSign" && cert.is_expired=false'
# M034 证书持有者apple.com但域名非apple.com
A["M034-S007"] = 'cert.subject="apple.com" && domain!="apple.com"'
# M035 证书全文含TrustAsia
A["M035-S010"] = 'cert="TrustAsia"'
# M036 证书持有者组织含Alibaba
A["M036-S007"] = 'cert.subject="Alibaba"'
# M037 sohu.com下, banner有内容且不含apache
A["M037-S008"] = 'domain="sohu.com" && banner!="" && banner!="apache"'
# M038 证书含adobe.com根域
A["M038-S009"] = 'cert="adobe.com"'
# M039 证书持有者根域qq.com
A["M039-S001"] = 'cert.subject="qq.com"'
# M040 证书含godaddy但根域非godaddy.com
A["M040-S011"] = 'cert="godaddy" && domain!="godaddy.com"'
# M041 证书含google.com, 排除蜜罐/欺诈
A["M041-S001"] = 'cert="google.com" && is_honeypot=false && is_fraud=false'
# M042 证书持有者含example.com或证书含example.com(含裸域与子域)
A["M042-S003"] = '(cert.subject="example.com" || cert="example.com")'
# M043 仍用TLS1.0且证书未过期
A["M043-S004"] = 'tls.version="TLS 1.0" && cert.is_expired=false'
# M044 JARM + 状态码307 + 响应头Content-Length: 0
A["M044-S007"] = 'jarm="21d21d00021d21d00021d21d21d21dce319294781f0e12867fd06896a000ea" && status_code="307" && header="Content-Length: 0"'
# M045 OS识别为linux
A["M045-S001"] = 'os="linux"'
# M046 中国大陆无绑定域名的网站类资产
A["M046-S013"] = 'country="CN" && is_domain=false'
# M047 证书颁发者与持有者不匹配(证书与域名不匹配)
A["M047-S011"] = 'cert.is_match=false'
# M048 加拿大, 证书与域名不匹配
A["M048-S001"] = 'country="CA" && cert.is_match=false'
# M049 土耳其 TLS JA3S指纹
A["M049-S010"] = 'country="TR" && tls.ja3s="f4febc55ea12b31ae17cfb7e614afda8"'
# M050 网络组织名称含Inc
A["M050-S004"] = 'org="Inc"'
# M051 美国OVH SAS, 2087端口, 标题完整WHM 登录
A["M051-S001"] = 'country="US" && org="OVH SAS" && port="2087" && title=="WHM 登录"'
# M052 .io域名下WordPress
A["M052-S002"] = 'host="*.io" && app="WordPress"'
# M053 厦门80端口且状态码200
A["M053-S009"] = 'city="Xiamen" && port="80" && status_code="200"'
# M054 英国DigitalOcean疑似/api/show
A["M054-S009"] = 'country="GB" && org="DigitalOcean" && body="/api/show"'
# M055 中国大陆Set-Cookie但正文/admin
A["M055-S003"] = 'country="CN" && header="Set-Cookie" && body="/admin"'
# M056 baidu.com中正文任一词或响应头Set-Cookie
A["M056-S008"] = 'domain="baidu.com" && (body="Login" || body="admin" || body="Dashboard" || body="管理" || header="Set-Cookie")'
# M057 13.0.0.0/8内网页: 200, 正文Welcome/Sign in/Dashboard任一
A["M057-S012"] = 'cidr="13.0.0.0/8" && status_code="200" && (body="Welcome" || body="Sign in" || body="Dashboard")'
# M058 德国与amazonaws.com相关, 五处任一命中
A["M058-S012"] = 'country="DE" && (host="amazonaws.com" || host="*.amazonaws.com" || domain="amazonaws.com" || cert="amazonaws.com")'
# M059 143端口证书序列1且含TAK等词
A["M059-S006"] = 'port="143" && (cert="TAK" || cert="ATAK" || cert="WinTAK" || cert="Squad" || cert="CoT")'
# M060 113.247.228.0/24 开放80/443/8080/8443任一
A["M060-S010"] = 'cidr="113.247.228.0/24" && (port="80" || port="443" || port="8080" || port="8443")'
# M061 标题Sign in或Login, 80端口, 非IPv6
A["M061-S005"] = '(title="Sign in" || title="Login") && port="80" && is_ipv6=false'
# M062 新西兰80或443且云服务商
A["M062-S010"] = 'country="NZ" && (port="80" || port="443") && is_cloud=true'
# M063 Banner含IIS和1. 或响应头含IIS
A["M063-S001"] = '(banner="IIS" && banner="1.") || header="IIS"'
# M064 日本已绑定域名网站:443,200,有效证书;标题或正文关键词
A["M064-S004"] = 'country="JP" && is_domain=true && port="443" && status_code="200" && cert.is_valid=true && ((title="大学" || title="学部" || title="研究") || (body="入学案内" || body="学生"))'
# M065 baidu.com主域: 标题词或子域词任一
A["M065-S004"] = 'domain="baidu.com" && ((title="系统" || title="平台" || title="登录" || title="管理" || title="API") || (host="api" || host="admin" || host="dev" || host="test" || host="uat" || host="pre"))'
# M066 邮件系统且标题含Mail, 非美国
A["M066-S006"] = 'app="邮件系统" && title="Mail" && country!="US"'
# M067 中国大陆8888端口, 排除港澳台
A["M067-S008"] = 'country="CN" && port="8888" && region!="HK" && region!="MO" && region!="TW"'
# M068 正文含Jenkins但标题不含
A["M068-S005"] = 'body="Jenkins" && title!="Jenkins"'
# M069 台湾AS20940 443 TLS1.3 排除Let's Encrypt/ZeroSSL
A["M069-S001"] = 'country="TW" && asn="20940" && port="443" && tls.version="TLS 1.3" && cert.issuer!="Let\'s Encrypt" && cert.issuer!="ZeroSSL"'
# M070 日本以外HFS, 200, 标题完整HFS /, 排除亚马逊云
A["M070-S003"] = 'country!="JP" && app="HFS" && status_code="200" && title=="HFS /" && cloud_name!="Amazon"'
# M071 不含港澳台的中国Grafana
A["M071-S006"] = 'country="CN" && app="Grafana" && region!="HK" && region!="MO" && region!="TW"'
# M072 .edu.ar高校, 80或443, 排除Nginx和WAF
A["M072-S008"] = 'host="*.edu.ar" && (port="80" || port="443") && server!="Nginx" && app!="WAF"'
# M073 IP 1.1.1.1 已绑定域名
A["M073-S003"] = 'ip="1.1.1.1" && is_domain=true'
# M074 证书根域aliyun.com, 有效, 资产无绑定域名
A["M074-S008"] = 'cert="aliyun.com" && cert.is_valid=true && is_domain=false'
# M075 中国已绑定域名, 响应头x-qiniu-zone, 主域非qiniudn.com
A["M075-S006"] = 'country="CN" && is_domain=true && header="x-qiniu-zone" && domain!="qiniudn.com"'
# M076 标题空, 200, fid与图标哈希非空, 域名含admin, 根域.cn或.com前4字符
A["M076-S004"] = 'title="" && status_code="200" && fid!="" && icon_hash!="" && host="admin" && (domain=".cn" || domain="*.com")'
# M077 源码含class="container"
A["M077-S009"] = 'body="class=\\"container\\""'
# M078 正文含Symfony异常片段(保留双引号与反斜杠)
A["M078-S009"] = 'body=\'"exception": "Symfony\\Component\\ErrorHandler\\Error\\FatalError"\''
# M079 40.88.14.254整个B段
A["M079-S006"] = 'cidr="40.88.0.0/16"'
# M080 47.96.0.1-47.96.0.5闭区间, 80或443
A["M080-S006"] = '(ip="47.96.0.1" || ip="47.96.0.2" || ip="47.96.0.3" || ip="47.96.0.4" || ip="47.96.0.5") && (port="80" || port="443")'
# M081 Grafana, 四项哈希至少两项非空 (programmatic)
fields = ["header_hash", "body_hash", "icon_hash", "fid"]
import itertools
combos = []
n = len(fields)
for k in range(2, n + 1):
    for c in itertools.combinations(fields, k):
        combos.append(" && ".join(f'{x}!=""' for x in c))
m081 = "app=\"Grafana\" && (" + " || ".join(combos) + ")"
A["M081-S011"] = m081
# M082 多主域(绑定或证书命中), 证书匹配且有效, 过期时间在窗口内
doms = ["kuaihuoyun.com", "wlhyos.cn", "wlhyos.com", "56ctms.com", "xiaokuaikeji.com", "zhihuiwuliu.org.cn"]
dom_clauses = " || ".join(f'domain="{d}"' for d in doms) + " || " + " || ".join(f'cert="{d}"' for d in doms)
A["M082-S001"] = f'({dom_clauses}) && cert.is_match=true && cert.is_valid=true && cert.notafter>"2024-01-01 00:00:00" && cert.notafter<"2026-08-30 18:17:49"'
# M083 Shodan转FOFA: Server: nginx + title Welcome
A["M083-S008"] = 'title="Welcome" && server="nginx"'
# M084 Google Dork转FOFA: intitle Grafana intext JavaScript
A["M084-S004"] = 'title="Grafana" && body="JavaScript"'
# M085 urlscan转FOFA: HTML含login, 域名含bet/casino/kasino
A["M085-S001"] = 'body="login" && (domain="bet" || domain="casino" || domain="kasino")'
# M086 Hunter转FOFA
A["M086-S006"] = 'domain="google" && domain!="google.com" && title=""'
# M087 Censys转FOFA
A["M087-S009"] = 'header="server: Tengine" || header="x-cache: MISS"'
# M088 修正: 香港, server精确Apache, 排除ASN4134
A["M088-S004"] = 'country="HK" && server=="Apache" && asn!="4134"'
# M089 修正: HFS 80全要; 8080仅200
A["M089-S010"] = '(app="HFS" && port="80") || (app="HFS" && port="8080" && status_code="200")'
# M090 依响应片段生成: Etag/Content-Length保留值, 其余字段名存在
A["M090-S008"] = 'header=\'Etag: "6aa522a3-493"\' && header="Content-Length: 1171" && header="Server" && header="Accept-Ranges" && header="Cache-Control" && header="Content-Type" && header="Date" && header="Last-Modified" && header="Vary"'
# M091 端口1000000非法 -> 不可转换
A["M091-S009"] = CANNOT
# M092 IP 1.2.3.4.5非法 -> 不可转换
A["M092-S007"] = CANNOT
# M093 国家码既等于US又不等于US -> 矛盾不可转换
A["M093-S004"] = CANNOT
# M094 最安全的网站 -> 模糊不可转换
A["M094-S004"] = CANNOT
# M095 容器内存使用率 -> 超出FOFA范围不可转换
A["M095-S010"] = CANNOT
# M096 Jenkins TCP Agent Listener 诊断字段名
A["M096-S001"] = '(body="Jenkins-Agent-Protocols" || body="Jenkins-Version" || body="Jenkins-Session" || body="Remoting-Minimum-Version" || banner="Jenkins-Agent-Protocols" || banner="Jenkins-Version" || banner="Jenkins-Session" || banner="Remoting-Minimum-Version")'
# M097 octodex.github.com 同图标
A["M097-S001"] = 'icon_hash="-698715305"'
# M098 与msp.gov.cn.poliec.com相同仿冒特征
A["M098-S001"] = 'is_fraud=true'
# M099 CVE-2021-41773 -> Apache httpd
A["M099-S001"] = 'app="Apache httpd"'
# M100 指定证书颁发者CN与序列号
A["M100-S001"] = 'cert.issuer="ZeroSSL ECC DV SSL CA 2" && cert="63:BD:83:71:62:14:83:6E:F8:00:75:35:61:61:EB:06"'

# ---- sanity: every 题号 from 题目.docx must be present ----
qdoc = docx.Document("/Users/xxx/code/test1/题目.docx")
ids = []
for p in qdoc.paragraphs:
    m = re.search(r'"题号":\s*"([^"]+)"', p.text)
    if m:
        ids.append(m.group(1))
missing = [i for i in ids if i not in A]
extra = [i for i in A if i not in ids]
print("题目数量:", len(ids), "答案数量:", len(A))
print("缺失:", missing)
print("多余:", extra)

# ---- fill the docx template ----
doc = docx.Document(TEMPLATE)
cur = None
filled = 0
for p in doc.paragraphs:
    m = re.search(r'"题号":\s*"([^"]+)"', p.text)
    if m:
        cur = m.group(1)
        continue
    if '"查询语句"' in p.text and cur:
        val = A.get(cur, "")
        # preserve leading whitespace and trailing comma
        lead = p.text[:len(p.text) - len(p.text.lstrip())]
        has_comma = p.text.rstrip().endswith(",")
        new_line = lead + '"查询语句": ' + json.dumps(val, ensure_ascii=False) + ("," if has_comma else "")
        # set run text
        if p.runs:
            p.runs[0].text = new_line
            for r in p.runs[1:]:
                r.text = ""
        else:
            p.add_run(new_line)
        filled += 1
        cur = None
print("已填充查询语句条目:", filled)

# set 选手名称
for p in doc.paragraphs:
    if '"选手名称"' in p.text:
        lead = p.text[:len(p.text) - len(p.text.lstrip())]
        has_comma = p.text.rstrip().endswith(",")
        new_line = lead + '"选手名称": ' + json.dumps(PLAYER, ensure_ascii=False) + ("," if has_comma else "")
        if p.runs:
            p.runs[0].text = new_line
            for r in p.runs[1:]:
                r.text = ""
        break

doc.save(OUT_DOCX)
print("已保存:", OUT_DOCX)

# ---- also produce valid 答案.json ----
answers = []
for i in ids:
    answers.append({"题号": i, "查询语句": A[i]})
out = {"选手名称": PLAYER, "参赛包编号": PKG, "答案": answers}
with open(OUT_JSON, "w", encoding="utf-8") as f:
    json.dump(out, f, ensure_ascii=False, indent=2)
print("已保存:", OUT_JSON)
