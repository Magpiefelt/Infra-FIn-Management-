// Shared layout and navigation components
import React from 'react'
import { Button } from '@/components/ui/button.jsx'
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from '@/components/ui/select.jsx'
import { User, Building } from 'lucide-react'
import { useAuth } from '../../hooks/index.js'
import albertaLogo from '../../assets/alberta-logo.png'

export function Header({ onRoleChange }) {
  const { currentUser, changeRole } = useAuth()

  const handleRoleChange = (newRole) => {
    changeRole(newRole)
    if (onRoleChange) {
      onRoleChange(newRole)
    }
  }

  return (
    <header className="bg-white border-b border-gray-200 px-6 py-4">
      <div className="flex items-center justify-between">
        <div className="flex items-center space-x-4">
          <img src={albertaLogo} alt="Alberta Government" className="h-8 w-auto" />
          <div>
            <h1 className="text-xl font-semibold text-gray-900">
              Project Financial Management Tool
            </h1>
            <p className="text-sm text-gray-600">Replacement System Demo</p>
          </div>
        </div>
        
        <div className="flex items-center space-x-4">
          <div className="flex items-center space-x-2">
            <User className="h-4 w-4 text-gray-500" />
            <span className="text-sm font-medium text-gray-700">{currentUser.name}</span>
          </div>
          
          <Select value={currentUser.role} onValueChange={handleRoleChange}>
            <SelectTrigger className="w-48">
              <SelectValue />
            </SelectTrigger>
            <SelectContent>
              <SelectItem value="vendor">
                <div className="flex items-center space-x-2">
                  <Building className="h-4 w-4" />
                  <span>Vendor/Contractor</span>
                </div>
              </SelectItem>
              <SelectItem value="pm">
                <div className="flex items-center space-x-2">
                  <User className="h-4 w-4" />
                  <span>Project Manager</span>
                </div>
              </SelectItem>
              <SelectItem value="spm">
                <div className="flex items-center space-x-2">
                  <User className="h-4 w-4" />
                  <span>Senior Project Manager</span>
                </div>
              </SelectItem>
              <SelectItem value="director">
                <div className="flex items-center space-x-2">
                  <User className="h-4 w-4" />
                  <span>Director</span>
                </div>
              </SelectItem>
              <SelectItem value="admin">
                <div className="flex items-center space-x-2">
                  <User className="h-4 w-4" />
                  <span>Administrative Staff</span>
                </div>
              </SelectItem>
            </SelectContent>
          </Select>
        </div>
      </div>
    </header>
  )
}

export function Layout({ children }) {
  return (
    <div className="min-h-screen bg-gray-50">
      {children}
    </div>
  )
}

export function PageHeader({ title, description, actions }) {
  return (
    <div className="bg-white border-b border-gray-200 px-6 py-4">
      <div className="flex items-center justify-between">
        <div>
          <h2 className="text-2xl font-bold text-gray-900">{title}</h2>
          {description && (
            <p className="text-sm text-gray-600 mt-1">{description}</p>
          )}
        </div>
        {actions && (
          <div className="flex items-center space-x-2">
            {actions}
          </div>
        )}
      </div>
    </div>
  )
}

export function LoadingSpinner({ size = 'md' }) {
  const sizeClasses = {
    sm: 'h-4 w-4',
    md: 'h-8 w-8',
    lg: 'h-12 w-12'
  }

  return (
    <div className="flex items-center justify-center">
      <div className={`animate-spin rounded-full border-2 border-gray-300 border-t-blue-600 ${sizeClasses[size]}`}></div>
    </div>
  )
}

export function ErrorMessage({ message, onRetry }) {
  return (
    <div className="bg-red-50 border border-red-200 rounded-md p-4">
      <div className="flex items-center justify-between">
        <p className="text-red-800">{message}</p>
        {onRetry && (
          <Button variant="outline" size="sm" onClick={onRetry}>
            Retry
          </Button>
        )}
      </div>
    </div>
  )
}

export function EmptyState({ title, description, action }) {
  return (
    <div className="text-center py-12">
      <h3 className="text-lg font-medium text-gray-900 mb-2">{title}</h3>
      <p className="text-gray-600 mb-4">{description}</p>
      {action}
    </div>
  )
}

