import { defineConfig } from 'vite';

/**
 * data-cleaner (공공데이터 결측치 정리 실습기) - 개발/빌드 설정
 *
 * 이 설정 파일은 앱 폴더 안에 있고, 웹서버의 문서 루트도 이 폴더다.
 *   npm start  →  http://localhost:8081/ 에서 이 폴더의 index.html 이 메인 페이지로 열린다.
 *
 * 포트를 8081 로 쓰는 이유: 형제 앱 binary-converter 가 8080 을 쓴다.
 * 두 앱을 동시에 띄워도 겹치지 않게 앱마다 포트를 다르게 준다.
 *
 * 설치·빌드 결과물(node_modules/, dist/)도 모두 이 폴더 안에 생긴다.
 * 프로젝트 루트(C:\project_AI)에는 아무것도 설치하지 않는다.
 */
export default defineConfig({
  // 문서 루트 = 이 설정 파일이 있는 폴더 (= data-cleaner)
  root: '.',

  // 빌드 결과물을 파일로 직접 열어도(file://) 경로가 깨지지 않도록 상대 경로로 빌드한다
  base: './',

  server: {
    port: 8081,
    strictPort: true,   // 8081이 사용 중이면 조용히 다른 포트로 옮기지 않고 알려 준다
    open: false         // 교실에서 자동으로 브라우저를 띄우려면 true 로 바꾼다
  },

  preview: {
    port: 8081,
    strictPort: true
  },

  build: {
    // 결과물은 이 폴더 안의 dist 에 모은다
    outDir: 'dist',
    emptyOutDir: true,
    sourcemap: false,
    // 엑셀을 읽고 쓰는 SheetJS 가 커서 기본 경고선(500kB)을 넘는다.
    // 교실에서는 한 번만 받아 두면 되므로 문제가 아니다. 경고선만 올려 둔다.
    chunkSizeWarningLimit: 800
  }
});
