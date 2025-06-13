// Server startup file
import app from './app.js'
import { initializeDatabase } from './services/database.js'

const PORT = process.env.PORT || 3001

// Initialize database and start server
function startServer() {
  try {
    // Initialize the database
    initializeDatabase()
    console.log('Database initialized successfully')
    
    // Start the server
    app.listen(PORT, '0.0.0.0', () => {
      console.log(`PFMT Backend server running on port ${PORT}`)
      console.log(`Health check: http://localhost:${PORT}/api/health`)
      console.log(`Environment: ${process.env.NODE_ENV || 'development'}`)
    })
  } catch (error) {
    console.error('Failed to start server:', error)
    process.exit(1)
  }
}

// Handle graceful shutdown
process.on('SIGTERM', () => {
  console.log('SIGTERM received, shutting down gracefully')
  process.exit(0)
})

process.on('SIGINT', () => {
  console.log('SIGINT received, shutting down gracefully')
  process.exit(0)
})

// Start the server
startServer()

