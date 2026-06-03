# BOM 校验工具 - 部署指南

## 方式一：Vercel 部署（推荐，免费）

### 步骤

1. **GitHub 推送代码**
   ```bash
   cd /workspace/bom-check-tool
   git remote add origin <你的GitHub仓库地址>
   git branch -M main
   git push -u origin main
   ```

2. **注册/登录 Vercel**
   - 访问 https://vercel.com
   - 使用 GitHub 账号登录

3. **导入项目**
   - 点击 "Add New Project"
   - 选择 `bom-check-tool` 仓库
   - 点击 "Deploy"

4. **获取公网地址**
   - 部署完成后，Vercel 会提供一个类似 `https://bom-check-tool-xxx.vercel.app` 的地址
   - **把这个地址发给别人即可使用**

---

## 方式二：Railway 部署

1. 访问 https://railway.app
2. 关联 GitHub 仓库
3. 自动检测 Node.js 项目并部署
4. 获得公网 URL

---

## 注意事项

- 文件上传限制：Vercel Serverless 函数最大请求体 4.5MB
- 如果需要支持更大的 Excel 文件，建议使用 Railway（支持 10MB）
