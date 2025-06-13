// Utility functions for formatting and validation
export const formatCurrency = (value) => {
  if (value === null || value === undefined) return 'N/A'
  return new Intl.NumberFormat('en-CA', {
    style: 'currency',
    currency: 'CAD',
    minimumFractionDigits: 0,
    maximumFractionDigits: 0
  }).format(value)
}

export const formatDate = (dateString) => {
  if (!dateString) return 'N/A'
  return new Date(dateString).toLocaleDateString('en-CA')
}

export const formatDateTime = (dateString) => {
  if (!dateString) return 'N/A'
  return new Date(dateString).toLocaleString('en-CA')
}

export const calculateVariance = (target, actual) => {
  if (!target || !actual) return null
  return target - actual
}

export const getVariancePercentage = (target, actual) => {
  if (!target || !actual) return null
  return ((actual - target) / target) * 100
}

export const getStatusColor = (status) => {
  switch (status?.toLowerCase()) {
    case 'green':
    case 'on track':
    case 'completed':
      return 'text-green-600 bg-green-50'
    case 'yellow':
    case 'minor variance':
    case 'in-progress':
      return 'text-yellow-600 bg-yellow-50'
    case 'red':
    case 'major variance':
    case 'overdue':
      return 'text-red-600 bg-red-50'
    default:
      return 'text-gray-600 bg-gray-50'
  }
}

export const validateEmail = (email) => {
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/
  return emailRegex.test(email)
}

export const validateRequired = (value) => {
  return value !== null && value !== undefined && value !== ''
}

export const validateNumber = (value) => {
  return !isNaN(value) && isFinite(value)
}

export const validatePositiveNumber = (value) => {
  return validateNumber(value) && value > 0
}

export const debounce = (func, wait) => {
  let timeout
  return function executedFunction(...args) {
    const later = () => {
      clearTimeout(timeout)
      func(...args)
    }
    clearTimeout(timeout)
    timeout = setTimeout(later, wait)
  }
}

export const generateId = () => {
  return Math.random().toString(36).substr(2, 9)
}

export const deepClone = (obj) => {
  return JSON.parse(JSON.stringify(obj))
}

export const isEmpty = (value) => {
  if (value === null || value === undefined) return true
  if (typeof value === 'string') return value.trim() === ''
  if (Array.isArray(value)) return value.length === 0
  if (typeof value === 'object') return Object.keys(value).length === 0
  return false
}

export const capitalizeFirst = (str) => {
  if (!str) return ''
  return str.charAt(0).toUpperCase() + str.slice(1)
}

export const truncateText = (text, maxLength = 100) => {
  if (!text || text.length <= maxLength) return text
  return text.substring(0, maxLength) + '...'
}

