# Hermes for Office

Add-in cho Microsoft Word và Excel: chat với một LLM về tài liệu đang mở, xem trước
đề xuất thay đổi, và chỉ ghi vào tài liệu khi bạn bấm **Áp dụng**.

Dùng được với mọi API tương thích OpenAI — OpenAI, OpenRouter, Azure OpenAI, hoặc
router bạn tự host.

```
Word / Excel add-in ──▶ Local Gateway (Caddy, localhost:8643) ──▶ Provider /chat/completions
```

**Local Gateway** là Caddy chạy trên máy bạn. Nó vừa phục vụ bundle add-in, vừa
chuyển tiếp mọi lời gọi lên **Provider** kèm khóa API. Add-in không bao giờ gọi
thẳng Provider và không giữ khóa nào cả — xem [ADR-0001](./docs/adr/0001-local-gateway-is-the-only-path-to-the-provider.md)
và [ADR-0002](./docs/adr/0002-the-api-key-lives-in-the-caddyfile.md).

| Ứng dụng | Trạng thái | Thư mục |
|----------|------------|---------|
| Word — chat trong task pane + sửa tài liệu bằng AI | ✅ dùng được | [`word/`](./word) |
| Excel — chat trong task pane + hàm `=HERMES.*` | ✅ dùng được | [`excel/`](./excel) |

## Cài đặt

Cần: Microsoft 365 bản desktop, Node.js 18+, npm, và [Caddy](https://caddyserver.com/docs/install).

Cài dependency cho hai add-in trước khi chạy script cài đặt:

```bash
npm ci --prefix word
npm ci --prefix excel
```

```bash
# macOS
bash scripts/install.sh
# Windows PowerShell
.\scripts\install.ps1
```

Script sẽ build add-in rồi hỏi ba thứ. Nếu Caddy chưa có trong `PATH`, script chỉ
cảnh báo; cài Caddy trước khi chạy Gateway.

| Hỏi | Ví dụ | Ghi vào đâu |
|-----|-------|-------------|
| URL Provider | `https://openrouter.ai/api/v1` | `Caddyfile` |
| Khóa API | `sk-…` | `Caddyfile` (chmod 600, git-ignored) |
| Tên model | `anthropic/claude-sonnet-4` | `config.json` |

Tên model được giữ **nguyên văn** — nhập đúng như Provider của bạn gọi nó.

Sau đó chạy Local Gateway và để nguyên cửa sổ đó:

```bash
npm run serve
```

Cuối cùng sideload manifest trong Word và Excel: **Home → Get Add-ins → Manage My
Add-ins → Upload My Add-in**, chọn `word/dist/manifest.xml` cho Word và
`excel/dist/manifest.xml` cho Excel.

Đổi Provider, khóa hay model về sau: chạy lại `npm run setup`. Nó chỉ ghi đè khối
giữa hai dòng mốc `# >>> hermes-proxy >>>` trong `Caddyfile`, nên các chỉnh tay
khác của bạn trong file đó vẫn còn.

**Cài đặt có từ trước ADR-0005 cần chạy lại `npm run setup` một lần.** Khối giữa
hai mốc giờ còn chứa `handle_errors` — thứ giúp task pane phân biệt "Gateway
không với tới Provider" với "Provider trả về lỗi". Thiếu nó thì add-in vẫn chạy
bình thường, chỉ là một `502` sẽ luôn bị đọc thành lỗi của Provider.

Cài phi tương tác (script, cài lại hàng loạt):

```bash
HERMES_API_KEY=sk-… node scripts/setup.mjs \
  --provider=https://openrouter.ai/api/v1 --model=anthropic/claude-sonnet-4
```

Khóa chỉ truyền qua biến môi trường, không qua tham số dòng lệnh — argv thì mọi
tiến trình trên máy đều đọc được và nó lọt vào lịch sử shell.

## Cách hoạt động

1. Task pane đọc phần bạn đang chọn (hoặc bảng, hoặc cả tài liệu/sheet) cùng câu
   bạn gõ, gói lại thành một **Snapshot** có giới hạn kích thước.
2. Gửi hội thoại tới Local Gateway; Gateway chuyển tiếp lên Provider kèm khóa.
3. Câu trả lời được tách thành phần văn xuôi và một **Proposal** — tập thay đổi
   được hiển thị thành thẻ. Lúc này chưa có gì được ghi vào tài liệu.
4. Bấm **Áp dụng** thì Office.js mới thực thi. Word có thể tô đỏ phần vừa sửa để
   bạn thấy ngay chỗ nào do AI thay đổi.

Excel có thêm các hàm gọi thẳng từ ô, không qua task pane: `=HERMES.CLASSIFY`,
`=HERMES.EXTRACT`, `=HERMES.SUMMARIZE`, `=HERMES.FORMULA_HELP`.

Các hàm này **mặc định tắt** — vì chúng nằm ngoài nút Áp dụng và một file bảng
tính chỉ cần *chứa sẵn* công thức là chúng tự chạy lúc mở. Bật khi bạn tin bảng
tính của mình, bằng cách thêm vào `config.json` rồi tải lại Excel:

```json
{ "name": "…", "model": "…", "customFunctions": true }
```

Xem [ADR-0003](./docs/adr/0003-custom-functions-are-off-by-default.md). `npm run
setup` giữ nguyên lựa chọn này khi bạn đổi Provider.

## Khi gọi bị lỗi

Một lời gọi đi qua **hai chặng**, và chúng hỏng vì những lý do khác nhau:

```
Task Pane ──Local Hop──▶ Local Gateway ──Upstream Hop──▶ Provider
```

Bấm biểu tượng **Kiểm tra kết nối** trên đầu task pane để thử riêng từng chặng —
không tốn token nào, và không cần phải gõ một câu hỏi thật để kích lỗi ra. Kết
quả là hai dòng, ví dụ:

```
Local Hop: OK
Upstream Hop: Local Gateway không kết nối được tới Provider. Kiểm tra mạng, rồi bấm Hỏi lại.
```

Task pane **không bao giờ hiện nguyên văn lỗi của Provider** — không body phản
hồi, không stack, không đường dẫn. Thay vào đó nó nói đúng một trong các trạng
thái đã đặt tên sẵn, mỗi trạng thái kèm đúng một việc cần làm (Gateway chưa
chạy → `npm run serve`; khóa bị từ chối → `npm run setup`; sai tên model → sửa
tên model; v.v.). Muốn xem lỗi gốc thì mở console của task pane, hoặc log Caddy ở
`~/Library/Logs/hermes-office/caddy.error.log`. Xem
[ADR-0005](./docs/adr/0005-disclosed-failures-are-a-fixed-set.md).

## Chạy Gateway tự động (tùy chọn, macOS)

Nếu không muốn mở terminal mỗi lần dùng add-in, bạn có thể để `launchd` giữ
Gateway chạy sẵn bằng một LaunchAgent ở `~/Library/LaunchAgents/com.hermes.caddy.plist`
với `RunAtLoad` + `KeepAlive`.

`scripts/install.sh` sẽ **hỏi** có bật hay không, **mặc định là Không** — vì đánh
đổi bảo mật bên dưới. Bật/tắt bất cứ lúc nào:

```bash
node scripts/launchagent.mjs --install   # cài và nạp ngay
node scripts/launchagent.mjs --remove    # gỡ hẳn
```

Cài không cần hỏi (dùng khi triển khai hàng loạt): đặt `HERMES_AUTOSTART=1` hoặc
`HERMES_AUTOSTART=0` trước khi chạy `install.sh`. Nếu plist đã tồn tại, script
giữ nguyên file của bạn trừ khi thêm `--force`. Windows chưa có cơ chế tương
đương — `install.ps1` không cài tự khởi động.

Khi có LaunchAgent, cách vận hành đổi khác — đọc kỹ:

| Việc | Lệnh |
|------|------|
| Nạp lại sau khi sửa `Caddyfile` | `launchctl kickstart -k gui/$(id -u)/com.hermes.caddy` |
| Dừng hẳn (tới lần đăng nhập sau) | `launchctl bootout gui/$(id -u)/com.hermes.caddy` |
| Xem log | `~/Library/Logs/hermes-office/caddy.error.log` |

- `npm run setup` **tự kickstart** job sau khi ghi `Caddyfile`. Không có bước này
  thì đổi Provider/khóa/model xong vẫn chạy cấu hình cũ: caddy giữ Caddyfile
  trong bộ nhớ, không đọc lại từ đĩa.
- `npm run serve` phát hiện cổng đã bị chiếm và **không** dựng instance thứ hai.
  Hai caddy cùng bind một cổng thì request bị chia cho cả hai — cái cũ vẫn chạy
  config cũ, triệu chứng là 401 lúc có lúc không.
- `npm run stop` chỉ bảo tiến trình thoát; `KeepAlive` bật lại ngay. Script sẽ
  kiểm tra lại và nói thẳng điều đó thay vì báo "đã dừng".

## Gỡ cài đặt

```bash
npm run stop     # dừng Local Gateway
```

Nếu dùng LaunchAgent thì `launchctl bootout gui/$(id -u)/com.hermes.caddy`, và xóa
`~/Library/LaunchAgents/com.hermes.caddy.plist` nếu muốn bỏ hẳn.

Xóa add-in trong Office qua **Home → Get Add-ins → Manage My Add-ins → Delete**.
Xóa `Caddyfile` để gỡ khóa API khỏi máy.

## Bảo mật

**Khóa API nằm trong `Caddyfile`, không nằm trong `config.json`.** `config.json`
được Gateway phục vụ qua HTTP nên bất kỳ tiến trình nào trên máy cũng đọc được;
vì vậy nó chỉ chứa `{name, model, customFunctions}` — không có gì bí mật. `Caddyfile` không bao giờ được phục vụ, được
git-ignore, và đặt quyền 600.

Đánh đổi đi kèm, cần biết rõ: **trong lúc Gateway đang chạy, mọi tiến trình local
đều có thể POST tới `https://localhost:8643/v1/*` mà không cần khóa** và tiêu
quota của bạn. Nó không lấy được khóa để mang đi nơi khác, nhưng nó dùng được
Provider của bạn. Nên:

- Chỉ chạy `npm run serve` khi đang thực sự dùng add-in; xong thì `npm run stop`.
- **Nếu bạn bật LaunchAgent tự chạy, biện pháp trên không còn**: Gateway lên từ
  lúc đăng nhập và tự bật lại khi bị dừng, tức là cửa sổ phơi nhiễm là toàn bộ
  phiên làm việc. Đó là đánh đổi có ý thức giữa tiện lợi và rủi ro — nếu không
  chấp nhận được, đừng cài LaunchAgent và chạy `npm run serve` thủ công.
- Đừng để nó chạy trên máy dùng chung hay máy không tin cậy.
- `Caddyfile` và `config.json` đều đã git-ignored. Đừng commit, đừng gửi cho ai.

Local Gateway cố tình **không** đặt header `Access-Control-Allow-Origin`. Task
pane cùng origin với nó nên không cần CORS; thêm wildcard vào sẽ cho phép bất kỳ
trang web nào mở trong trình duyệt trên máy này điều khiển Provider của bạn.

### Prompt injection

Nội dung tài liệu và bảng tính là **đầu vào không đáng tin** — chữ trong file bạn
mở có thể tác động đến điều model đề xuất. Hai lớp phòng vệ, cả hai đều quan trọng:

- **Không gì được ghi khi bạn chưa duyệt.** Mọi thay đổi đều hiện thành thẻ và chỉ
  chạy khi bấm Áp dụng. Hãy đọc thẻ — đó chính là ranh giới bảo mật.
- **Thẻ hiện đúng thứ sắp được ghi.** Một hành động ghi cả vùng (`setCells`) hiện
  luôn nội dung của nó, không phải mỗi dòng "500 rows"; một hành động thay thế
  trong Word hiện số chỗ sẽ bị đổi *trước* khi bạn bấm. Ranh giới chỉ có giá trị
  bằng đúng những gì bạn đọc được trên thẻ.
- **Đề xuất cũ không ghi đè tài liệu mới.** Nếu sheet đã thay đổi kể từ lúc tạo đề
  xuất, Excel từ chối áp dụng thay vì ghi mù — thẻ bạn duyệt không còn mô tả đúng
  hiện trạng nữa.
- **Excel không bao giờ áp dụng giá trị model đề xuất thành công thức sống.** Chuỗi
  bắt đầu bằng `=`, `+`, `-` hoặc `@` bị ép thành text (`literalCellValue` trong
  [`excel/src/taskpane/actions.js`](./excel/src/taskpane/actions.js)), nên một
  `=WEBSERVICE(...)` được cấy vào không thể biến thành kênh rò rỉ dữ liệu.
- **Hàm `=HERMES.*` mặc định tắt.** Đây là đường duy nhất mà nội dung file tự nó
  chạm tới Provider được, không qua nút Áp dụng nào — nên nó phải được bật bằng
  tay trong `config.json`.

## Phát triển

```bash
npm test         # bộ test node:test (không cần cài dependency)
npm run lint     # lint cả hai add-in
npm run build    # build word/ rồi excel/
npm run setup    # cấu hình lại Provider
npm run serve    # chạy Local Gateway
npm run stop     # dừng Local Gateway
```

Cần Node 18+ (`.nvmrc` ghim 22). Sau khi cài dependency cho `word/` và `excel/`,
CI chạy test trên 18/20/22, lint, build, và kiểm
tra rằng bundle build ra không chứa chuỗi nào giống khóa API.

Từ vựng dùng chung trong code và tài liệu nằm ở [`CONTEXT.md`](./CONTEXT.md);
các quyết định kiến trúc nằm ở [`docs/adr/`](./docs/adr).

## Bố cục repo

```
hermes-office/
├── word/                 # add-in Word (Office.js, webpack)
│   └── src/taskpane/      # chat UI, selection-mgr, proposal-mgr
├── excel/                # add-in Excel
│   └── src/taskpane/      # chat UI, actions.js (parse + literalise)
├── shared/                # code dùng chung, import bằng đường dẫn tương đối
│   ├── hermes.js           # client gọi Local Gateway (timeout + 1 lần thử lại)
│   ├── parsers.js          # chuyển đổi cột, tách JSON, chữ ký snapshot
│   └── proposal-card.js    # UI thẻ đề xuất, toast, thanh ngữ cảnh
├── scripts/               # setup / serve / stop / install
├── Caddyfile.example      # mẫu Local Gateway; setup.mjs sinh ra Caddyfile
└── config.example.json    # mẫu {name, model, customFunctions}
```

Hai add-in không nằm trong npm workspace; `shared/` được dùng chung thuần bằng
import tương đối (`word/src/shared/hermes.js` và `excel/src/shared/hermes.js` đều
`export * from "../../../shared/hermes.js"`).

## Giấy phép

MIT — xem [LICENSE](./LICENSE).
