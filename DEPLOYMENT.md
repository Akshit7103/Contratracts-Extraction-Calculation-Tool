# Render Deployment Guide

## Prerequisites
1. GitHub account
2. Render account (sign up at render.com)

## Step-by-Step Deployment

### 1. Initialize Git Repository
```bash
git init
git add .
git commit -m "Initial commit - AMEX Incentive Calculator"
```

### 2. Create GitHub Repository
1. Go to GitHub and create a new repository named `amex-incentive-calculator`
2. Don't initialize with README (we already have files)
3. Copy the repository URL

### 3. Push to GitHub
```bash
git remote add origin https://github.com/YOUR_USERNAME/amex-incentive-calculator.git
git branch -M main
git push -u origin main
```

### 4. Deploy to Render
1. Log in to Render Dashboard
2. Click "New" → "Web Service"
3. Connect your GitHub repository
4. Select the `amex-incentive-calculator` repository
5. Configure deployment:
   - **Name**: `amex-incentive-calculator`
   - **Region**: Choose closest to your users
   - **Branch**: `main`
   - **Root Directory**: (leave blank)
   - **Runtime**: `Python 3`
   - **Build Command**: (auto-detected from render.yaml)
   - **Start Command**: (auto-detected from render.yaml)

### 5. Environment Variables
Render will automatically use the environment variables defined in `render.yaml`:
- `FLASK_ENV=production`
- `FLASK_SECRET_KEY` (auto-generated)
- `PYTHONPATH=.`

### 6. Deploy
1. Click "Create Web Service"
2. Wait for deployment to complete (5-10 minutes)
3. Access your app at the provided Render URL

## Features Enabled
✅ Auto file type detection  
✅ Enhanced document processing  
✅ Contract term extraction  
✅ CV Tier and NACV calculations  
✅ CHD adjustments  
✅ Persistent file storage  
✅ Health check endpoint  

## Troubleshooting

### Common Issues:
1. **Build fails on spaCy model download**
   - Render automatically downloads `en_core_web_sm` during build

2. **Enhanced processing not working**
   - Check logs for missing dependencies
   - All required packages are in requirements.txt

3. **File uploads not working**
   - Render provides 1GB persistent disk mounted at `/uploads`

4. **Memory issues**
   - Free tier has 512MB RAM limit
   - Consider upgrading if processing large files

### Logs Access:
- Go to Render Dashboard → Your Service → Logs
- Monitor real-time deployment and application logs

## Production Notes
- Maximum file upload: 16MB
- Timeout: 120 seconds for processing
- Health check: `/healthz` endpoint
- Auto-scaling: Disabled on free tier
- HTTPS: Automatically enabled

## Updating the App
```bash
git add .
git commit -m "Update description"
git push origin main
```
Render will automatically redeploy when you push to the main branch.