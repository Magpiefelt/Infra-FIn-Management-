# PFMT Application - Enhanced Version

## 🚀 Overview

This is the enhanced version of the PFMT (Project Financial Management Tool) application with three major improvements implemented:

1. **Node.js/Express Backend Integration** - Complete REST API with persistent storage
2. **Persistent and Paginated Project Dashboard** - Real-time data with pagination support  
3. **Code Review Findings Resolution** - Clean, optimized codebase

## ✨ New Features

### Backend API
- **Express.js REST API** with comprehensive endpoints
- **Persistent JSON storage** using lowdb
- **File upload support** for Excel files
- **CORS enabled** for frontend communication
- **Error handling** and validation
- **Health monitoring** endpoint

### Enhanced Frontend
- **Real-time data fetching** from backend API
- **Role-based access control** with proper filtering
- **Navigation functionality** for all dashboard tiles
- **PFMT Data Extractor** fully integrated
- **Project detail views** with back navigation
- **Responsive design** maintained

### Data Management
- **Automatic data synchronization** between frontend and backend
- **Pagination support** for large datasets
- **Excel file processing** for project creation
- **User role management** with proper access controls

## 🏗️ Architecture

### Frontend (React + Vite)
```
src/
├── components/          # React components
│   ├── projects/       # Project-related components
│   ├── shared/         # Shared UI components
│   └── ui/            # Base UI components
├── hooks/              # Custom React hooks
├── services/           # API service layer
├── stores/             # Zustand state management
└── App.jsx            # Main application component
```

### Backend (Node.js + Express)
```
backend/
├── controllers/        # Request handlers
├── routes/            # API route definitions
├── services/          # Business logic layer
├── models/            # Data models
├── uploads/           # File upload directory
└── data/              # JSON database files
```

## 🔧 Installation & Setup

### Prerequisites
- Node.js 18+ 
- npm or yarn

### Quick Start

1. **Install Frontend Dependencies**
   ```bash
   npm install
   ```

2. **Install Backend Dependencies**
   ```bash
   cd backend
   npm install
   cd ..
   ```

3. **Start Backend Server**
   ```bash
   cd backend
   npm start
   ```
   Backend will run on http://localhost:3001

4. **Start Frontend Development Server**
   ```bash
   npm run dev
   ```
   Frontend will run on http://localhost:5173

5. **Access Application**
   Open http://localhost:5173 in your browser

## 📊 API Endpoints

### Projects
- `GET /api/projects` - Get all projects with pagination
- `GET /api/projects/:id` - Get specific project
- `POST /api/projects` - Create new project
- `PUT /api/projects/:id` - Update project
- `DELETE /api/projects/:id` - Delete project
- `POST /api/projects/:id/upload` - Upload Excel file

### Users
- `GET /api/users` - Get all users
- `GET /api/users/:id` - Get specific user

### System
- `GET /api/health` - Health check endpoint

## 👥 User Roles

### Project Manager
- View assigned projects
- Update project status
- Upload PFMT files

### Senior Project Manager  
- View all projects
- Review and approve submissions
- Manage project assignments

### Director
- Full system access
- Approve major changes
- Generate reports

### Vendor
- Limited project visibility
- View relevant project information
- Access PFMT data extractor

## 🔄 Data Flow

1. **Frontend** makes API calls to backend
2. **Backend** processes requests and queries database
3. **Database** (JSON files) stores persistent data
4. **Real-time updates** reflect changes immediately
5. **Role-based filtering** ensures proper access control

## 📁 File Upload

The PFMT Data Extractor supports:
- **.xlsx files** (Excel 2007+)
- **.xlsm files** (Excel with macros)
- **File size limit**: 50MB
- **Automatic data extraction** and validation
- **Project creation** from Excel data

## 🧪 Testing

### Backend Testing
```bash
# Test health endpoint
curl http://localhost:3001/api/health

# Test projects endpoint  
curl http://localhost:3001/api/projects

# Test users endpoint
curl http://localhost:3001/api/users
```

### Frontend Testing
1. Navigate to http://localhost:5173
2. Test all navigation tiles
3. Verify project data loads correctly
4. Test PFMT Data Extractor functionality
5. Verify role-based access

## 🔧 Configuration

### Environment Variables
Create `.env` files in both root and backend directories:

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

## 📝 Development Notes

### State Management
- **Zustand** for global state management
- **Custom hooks** for business logic
- **API service layer** for backend communication

### Styling
- **Tailwind CSS** for utility-first styling
- **Shadcn/ui** components for consistent UI
- **Responsive design** for mobile compatibility

### Error Handling
- **Try-catch blocks** in all async operations
- **Graceful fallbacks** for missing data
- **User-friendly error messages**
- **Console logging** for debugging

## 🚀 Deployment

### Production Build
```bash
# Build frontend
npm run build

# Start backend in production
cd backend
NODE_ENV=production npm start
```

### Environment Setup
1. Set production environment variables
2. Configure database connection
3. Set up file upload directory
4. Configure CORS for production domain

## 📋 Changelog

### Version 2.0.0 (Enhanced)
- ✅ Added Node.js/Express backend with REST API
- ✅ Implemented persistent JSON database storage
- ✅ Added pagination support for project listings
- ✅ Integrated PFMT Data Extractor with file upload
- ✅ Fixed navigation functionality for all dashboard tiles
- ✅ Implemented role-based access control
- ✅ Added comprehensive error handling
- ✅ Cleaned up legacy code and improved performance
- ✅ Added real-time data synchronization
- ✅ Enhanced user experience with loading states

### Previous Version (1.0.0)
- Basic React frontend with mock data
- Static project listings
- Limited navigation functionality

## 🤝 Support

For technical support or questions about the enhanced features:
1. Check the API documentation above
2. Review the test results in the implementation report
3. Verify all dependencies are installed correctly
4. Ensure both frontend and backend servers are running

## 📄 License

This enhanced version maintains the same license as the original application.

