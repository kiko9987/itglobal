import { defineConfig } from 'vite'
import { resolve } from 'path'
import { fileURLToPath, URL } from 'node:url'

// Flask 서버 포트 (환경변수 또는 기본값)
const FLASK_PORT = process.env.FLASK_PORT || process.env.PORT || 5000;
const FLASK_TARGET = `http://localhost:${FLASK_PORT}`;

export default defineConfig({
  // Flask static 폴더와 통합
  root: 'src',
  base: process.env.NODE_ENV === 'production' ? '/static/dist/' : '/',  // 개발 시에는 루트 경로 사용

  build: {
    outDir: '../static/dist',
    emptyOutDir: true,
    manifest: true,  // manifest.json 생성
    rollupOptions: {
      input: {
        // 각 페이지별 진입점
        'project-list': resolve(__dirname, 'src/js/pages/project-list.js'),
        'leads': resolve(__dirname, 'src/js/pages/leads.js'),
        'stats-page': resolve(__dirname, 'src/js/pages/stats-page.js'),
        // 전역 유틸리티
        'authInterceptor': resolve(__dirname, 'src/js/utils/authInterceptor.js'),
        // 공통 스타일
        'main': resolve(__dirname, 'src/css/main.css'),
        'stats': resolve(__dirname, 'src/css/pages/stats.css')
      },
      output: {
        // 파일명 패턴 설정
        entryFileNames: 'js/[name].[hash].js',
        chunkFileNames: 'js/[name].[hash].js',
        assetFileNames: (assetInfo) => {
          if (assetInfo.name.endsWith('.css')) {
            return 'css/[name].[hash].css'
          }
          return 'assets/[name].[hash][extname]'
        },
        // 수동 청크 분리 (Code Splitting 최적화)
        manualChunks: (id) => {
          // node_modules를 vendor 청크로 분리
          if (id.includes('node_modules')) {
            return 'vendor';
          }
          // 공통 컴포넌트 분리
          if (id.includes('/src/js/components/')) {
            return 'components';
          }
          // 유틸리티 분리
          if (id.includes('/src/js/utils/') && !id.includes('authInterceptor')) {
            return 'utils';
          }
        }
      }
    },

    // 소스맵 생성 (프로덕션에서는 비활성화)
    sourcemap: process.env.NODE_ENV !== 'production',

    // 번들 크기 분석
    reportCompressedSize: true,

    // 청크 크기 제한
    chunkSizeWarningLimit: 500,  // 500KB로 낮춤

    // 압축 최적화
    minify: 'terser',
    terserOptions: {
      compress: {
        drop_console: process.env.NODE_ENV === 'production',  // 프로덕션에서 console 제거
        drop_debugger: true,
        pure_funcs: ['console.log', 'console.info', 'console.debug']  // 특정 console 함수 제거
      }
    }
  },

  css: {
    postcss: {
      plugins: [
        {
          postcssPlugin: 'internal:charset-removal',
          AtRule: {
            charset: (atRule) => {
              if (atRule.name === 'charset') {
                atRule.remove();
              }
            }
          }
        }
      ]
    }
  },

  server: {
    host: '0.0.0.0', // IPv4와 IPv6 모두 바인딩
    port: 5173,
    strictPort: false, // 포트가 사용 중이면 자동으로 다음 포트 사용
    // Flask 개발 서버와 함께 사용
    proxy: {
      // API 요청들은 Flask 서버로 프록시
      '/api': FLASK_TARGET,
      '/auth': FLASK_TARGET,
      '/login': FLASK_TARGET,
      '/logout': FLASK_TARGET,
      '/css-monitor': FLASK_TARGET,
      // 페이지 요청들도 Flask 서버로 프록시
      '/projects': FLASK_TARGET,
      '/leads': FLASK_TARGET,
      '/stats': FLASK_TARGET,
      '/user': FLASK_TARGET,
      '/admin': FLASK_TARGET,
      '/docs-help': FLASK_TARGET,
      '/data': FLASK_TARGET,
      // Static 파일들도 Flask 서버로 프록시 (Vite가 처리하지 않는 것들)
      '/static': {
        target: FLASK_TARGET,
        changeOrigin: true,
        // Vite에서 처리하는 dist 파일들은 제외
        bypass: (req, res, options) => {
          if (req.url.includes('/static/dist/')) {
            return null; // Vite가 처리하도록 함
          }
        }
      },
      // 파비콘과 기타 정적 파일들
      '/favicon.ico': FLASK_TARGET,
      // Socket.IO WebSocket 프록시 설정
      '/socket.io': {
        target: FLASK_TARGET,
        changeOrigin: true,
        ws: true // WebSocket 지원 활성화
      }
    }
  },

  resolve: {
    alias: {
      '@': fileURLToPath(new URL('./src', import.meta.url))
    }
  }
})