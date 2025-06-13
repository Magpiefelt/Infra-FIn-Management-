# PFMT Application - Quick Start Guide

## Overview
The enhanced PFMT application now includes a Node.js/Express backend with persistent storage and pagination capabilities. This guide will help you quickly start and test the application.

## Prerequisites
- Node.js v20+ installed
- npm package manager
- Modern web browser

## Quick Start

### 1. Start the Backend Server
```bash
cd pfmt-replacement-demo/backend
npm install
npm start
```
The backend server will start on `http://localhost:3001`

### 2. Start the Frontend Application
```bash
cd pfmt-replacement-demo
npm install
npm run dev
```
The frontend application will start on `http://localhost:5174`

### 3. Access the Application
Open your web browser and navigate to `http://localhost:5174`

## Testing the Enhanced Features

### Backend API Testing
Test the API endpoints directly:

```bash
# Health check
curl http://localhost:3001/api/health

# Get all projects with pagination
curl http://localhost:3001/api/projects

# Get all users
curl http://localhost:3001/api/users
```

### Frontend Features
1. **Homepage**: View the dashboard with user information
2. **Projects Page**: Navigate to `/projects` to see the project list
3. **Role-Based Access**: Notice different content based on user role
4. **Error Handling**: Application gracefully handles errors

## Sample Data
The application comes pre-loaded with:
- 3 sample projects (Calgary Elementary, Red Deer Community Center, Edmonton Hospital)
- 5 sample users with different roles
- Realistic project data including budgets, timelines, and status

## Key Features Implemented
✅ **Backend Integration**: Complete REST API with Express.js  
✅ **Persistent Storage**: JSON-based database with lowdb  
✅ **Pagination**: Page-based navigation for large datasets  
✅ **Code Quality**: Cleaned up legacy code and improved structure  
✅ **Error Handling**: Comprehensive error boundaries and fallbacks  

## Known Issues & Next Steps
- Store integration needs completion for automatic data fetching
- Role mapping between frontend/backend needs alignment
- Navigation tiles need click handlers

## Support
For technical issues or questions, refer to the detailed Implementation Report (implementation-report.pdf) which contains comprehensive documentation of all changes and recommendations.

---
**Quick Start Guide** | **PFMT Enhancement Project** | **June 2025**

