// Vite chỉ để chụp ảnh tài liệu: gốc là ai-work (mã nhánh main), /api và /ws chuyển về AI Gateway giả lập 127.0.0.1:8098.
// localStorage "shot_web"=1: chạy như bản Web (API_GATEWAY_URL cố định = chính tên miền), không hiện mục chọn AI Gateway.
// localStorage "shot_view": màn hình xác thực ban đầu (FACEID...).
import { createRequire } from 'node:module';
const require = createRequire('/home/local/ai-work/package.json');
const react = require('@vitejs/plugin-react').default ?? require('@vitejs/plugin-react');

const docShots = {
  name: 'doc-shots',
  enforce: 'pre',
  transformIndexHtml: (html) => html.replace('</head>',
    '<script>try { if (localStorage.getItem("shot_web")) window.TARA_CONFIG.API_GATEWAY_URL = location.origin; } catch (e) {}</script>\n  </head>'),
  transform(code, id) {
    if (!id.endsWith('/src/App.jsx')) return null;
    return code.replace("useState('LOGIN')", "useState(() => localStorage.getItem('shot_view') || 'LOGIN')")
      .replace('\n  const startWebcamHardware = async () => {',
        "\n  useEffect(() => { if (localStorage.getItem('shot_view') === 'FACEID') setTimeout(() => startWebcamHardware(), 500); }, []);\n  const startWebcamHardware = async () => {");
  },
};

export default {
  root: '/home/local/ai-work',
  configFile: false,
  plugins: [react(), docShots],
  server: {
    port: 5175, strictPort: true, host: '127.0.0.1',
    proxy: { '/api': 'http://127.0.0.1:8098', '/ws': { target: 'ws://127.0.0.1:8098', ws: true } },
  },
};
