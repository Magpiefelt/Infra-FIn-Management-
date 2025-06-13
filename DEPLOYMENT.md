# PFMT Application - Deployment Guide

## 🚀 Quick Deployment

### Option 1: Local Development
```bash
# 1. Install dependencies
npm install
cd backend && npm install && cd ..

# 2. Start backend
cd backend && npm start &

# 3. Start frontend  
npm run dev

# 4. Access application
# Frontend: http://localhost:5173
# Backend API: http://localhost:3001
```

### Option 2: Production Build
```bash
# 1. Build frontend
npm run build

# 2. Serve built files
npm run preview

# 3. Start backend in production
cd backend && NODE_ENV=production npm start
```

## 📋 Pre-deployment Checklist

- [ ] Node.js 18+ installed
- [ ] All dependencies installed (`npm install`)
- [ ] Backend dependencies installed (`cd backend && npm install`)
- [ ] Environment variables configured (if needed)
- [ ] Ports 3001 and 5173 available
- [ ] File upload directory writable

## 🔧 Configuration

### Default Ports
- **Frontend**: 5173 (development) / 4173 (preview)
- **Backend**: 3001

### Environment Variables (Optional)
Create `.env` files if you need to customize:

**Frontend (.env)**
```
VITE_API_BASE_URL=http://localhost:3001/api
```

**Backend (backend/.env)**  
```
NODE_ENV=development
PORT=3001
CORS_ORIGIN=http://localhost:5173
```

## 🧪 Verification Steps

After deployment, verify these features work:

1. **Home Page Navigation**
   - [ ] All tiles clickable and navigate correctly
   - [ ] "All Projects" → Projects page
   - [ ] "PFMT Library" → Data Extractor

2. **Projects Page**
   - [ ] Projects load from backend API
   - [ ] All 3 projects display with complete data
   - [ ] "View Details" buttons work
   - [ ] Back navigation works

3. **PFMT Data Extractor**
   - [ ] Upload interface displays
   - [ ] File selection works
   - [ ] Supports .xlsx and .xlsm files

4. **Backend API**
   - [ ] Health check: `curl http://localhost:3001/api/health`
   - [ ] Projects: `curl http://localhost:3001/api/projects`
   - [ ] Users: `curl http://localhost:3001/api/users`

## 🐛 Troubleshooting

### Frontend Issues
- **Blank page**: Check browser console for errors
- **API errors**: Verify backend is running on port 3001
- **Navigation not working**: Clear browser cache and refresh

### Backend Issues  
- **Port in use**: Change PORT in backend/.env or kill existing process
- **CORS errors**: Verify CORS_ORIGIN matches frontend URL
- **File upload fails**: Check uploads/ directory permissions

### Common Solutions
```bash
# Kill processes on ports
lsof -ti:3001 | xargs kill -9
lsof -ti:5173 | xargs kill -9

# Restart with clean install
rm -rf node_modules backend/node_modules
npm install && cd backend && npm install && cd ..
```

## 📊 Performance Notes

- **Initial load**: ~2-3 seconds for full application
- **API response time**: <100ms for typical requests  
- **File upload**: Supports up to 50MB Excel files
- **Concurrent users**: Tested with multiple browser sessions

## 🔄 Updates & Maintenance

### Updating the Application
1. Replace the entire application directory
2. Run `npm install` in both root and backend
3. Restart both frontend and backend servers
4. Verify all functionality works

### Database Maintenance
- Database files stored in `backend/data/`
- Automatic backups recommended for production
- JSON format allows easy inspection and editing

## 📞 Support

If you encounter issues:
1. Check this deployment guide
2. Review the main README.md
3. Check browser console and terminal output
4. Verify all prerequisites are met

## ✅ Success Indicators

You'll know the deployment is successful when:
- ✅ Home page loads with navigation tiles
- ✅ Projects page shows 3 projects with data
- ✅ PFMT Data Extractor is accessible
- ✅ All navigation works smoothly
- ✅ Backend API responds to health checks
- ✅ No console errors in browser

