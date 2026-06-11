import './globals.css';

export const metadata = {
  title: '数据可视化分析工具',
  description: '专业的数据可视化分析平台'
};

export default function RootLayout({ children }) {
  return (
    <html lang="zh-CN">
      <body>{children}</body>
    </html>
  );
}