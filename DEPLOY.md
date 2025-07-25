# Deployment Guide

This guide covers deploying the Excel to LLM Converter to various platforms.

## Prerequisites

- Node.js 18+ installed locally
- Git repository with your code
- Account on your chosen deployment platform

## Local Build

First, ensure your project builds successfully:

```bash
# Install dependencies
npm install

# Run build
npm run build

# Test the production build
npm run preview
```

## Deployment Options

### 1. Vercel (Recommended)

#### Option A: Deploy with Git

1. Push your code to GitHub/GitLab/Bitbucket
2. Visit [vercel.com](https://vercel.com)
3. Click "New Project"
4. Import your Git repository
5. Vercel will auto-detect Vite configuration
6. Click "Deploy"

#### Option B: Deploy with CLI

```bash
# Install Vercel CLI
npm i -g vercel

# Deploy
vercel

# Follow the prompts
```

### 2. Netlify

#### Option A: Drag & Drop

1. Run `npm run build`
2. Visit [app.netlify.com](https://app.netlify.com)
3. Drag the `dist` folder to the deployment area

#### Option B: Git Integration

1. Push to GitHub/GitLab/Bitbucket
2. In Netlify, click "New site from Git"
3. Choose your repository
4. Build settings are auto-configured via `netlify.toml`
5. Click "Deploy site"

#### Option C: CLI Deployment

```bash
# Install Netlify CLI
npm i -g netlify-cli

# Build and deploy
npm run build
netlify deploy --prod --dir=dist
```

### 3. GitHub Pages

1. Install gh-pages:
   ```bash
   npm install --save-dev gh-pages
   ```

2. Add to `package.json`:
   ```json
   {
     "scripts": {
       "predeploy": "npm run build",
       "deploy": "gh-pages -d dist"
     }
   }
   ```

3. Update `vite.config.ts`:
   ```typescript
   export default defineConfig({
     base: '/your-repo-name/',
     // ... rest of config
   })
   ```

4. Deploy:
   ```bash
   npm run deploy
   ```

### 4. AWS S3 + CloudFront

1. Build the project:
   ```bash
   npm run build
   ```

2. Create S3 bucket with static website hosting enabled

3. Upload `dist` folder contents to S3

4. Create CloudFront distribution pointing to S3

5. Update bucket policy for CloudFront access

### 5. Docker Deployment

Create a `Dockerfile`:

```dockerfile
# Build stage
FROM node:18-alpine as build
WORKDIR /app
COPY package*.json ./
RUN npm ci
COPY . .
RUN npm run build

# Production stage
FROM nginx:alpine
COPY --from=build /app/dist /usr/share/nginx/html
COPY nginx.conf /etc/nginx/conf.d/default.conf
EXPOSE 80
CMD ["nginx", "-g", "daemon off;"]
```

Create `nginx.conf`:

```nginx
server {
    listen 80;
    location / {
        root /usr/share/nginx/html;
        index index.html;
        try_files $uri $uri/ /index.html;
    }
}
```

Build and run:
```bash
docker build -t excel-llm-converter .
docker run -p 8080:80 excel-llm-converter
```

## Environment Variables

This app runs entirely in the browser and doesn't require environment variables. However, if you need to add API endpoints later:

1. Create `.env.production`:
   ```
   VITE_API_URL=https://your-api.com
   ```

2. Access in code:
   ```typescript
   const apiUrl = import.meta.env.VITE_API_URL
   ```

## Post-Deployment Checklist

- [ ] Test file upload functionality
- [ ] Verify all output formats work
- [ ] Check responsive design on mobile
- [ ] Test dark mode
- [ ] Verify formula extraction
- [ ] Check download functionality
- [ ] Test with large Excel files
- [ ] Verify HTTPS is enabled
- [ ] Check browser console for errors

## Performance Optimization

The build is already optimized with:
- Code splitting for xlsx library
- Minification with Terser
- Tree shaking
- Compressed assets

## Monitoring

Consider adding:
- Google Analytics
- Error tracking (Sentry)
- Performance monitoring

## Troubleshooting

### Build Fails
- Clear `node_modules` and reinstall: `rm -rf node_modules && npm install`
- Check Node.js version: `node --version` (should be 18+)

### Deployment Fails
- Check build logs for errors
- Ensure all dependencies are in `package.json`
- Verify deployment platform settings

### App Not Loading
- Check browser console for errors
- Verify all assets are loading (Network tab)
- Check for CORS issues if using external resources

## Security Notes

- All processing happens client-side
- No data is sent to servers
- No cookies or tracking by default
- Content Security Policy headers configured in deployment files