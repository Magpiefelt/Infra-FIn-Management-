// Tests for utility functions
import { describe, it, expect } from 'vitest'
import {
  formatCurrency,
  formatDate,
  calculateVariance,
  getVariancePercentage,
  validateEmail,
  validateRequired,
  validateNumber,
  validatePositiveNumber,
  isEmpty,
  capitalizeFirst,
  truncateText
} from '../utils/index.js'

describe('Utility Functions', () => {
  describe('formatCurrency', () => {
    it('should format positive numbers as Canadian currency', () => {
      expect(formatCurrency(1000)).toBe('$1,000')
      expect(formatCurrency(1234567)).toBe('$1,234,567')
    })

    it('should handle zero and negative numbers', () => {
      expect(formatCurrency(0)).toBe('$0')
      expect(formatCurrency(-1000)).toBe('-$1,000')
    })

    it('should return N/A for null or undefined', () => {
      expect(formatCurrency(null)).toBe('N/A')
      expect(formatCurrency(undefined)).toBe('N/A')
    })
  })

  describe('formatDate', () => {
    it('should format valid date strings', () => {
      // Note: formatDate uses toLocaleDateString which may vary by timezone
      const result = formatDate('2024-01-15')
      expect(result).toMatch(/2024-01-1[45]/) // Allow for timezone differences
    })

    it('should return N/A for invalid dates', () => {
      expect(formatDate(null)).toBe('N/A')
      expect(formatDate('')).toBe('N/A')
      expect(formatDate(undefined)).toBe('N/A')
    })
  })

  describe('calculateVariance', () => {
    it('should calculate variance correctly', () => {
      expect(calculateVariance(100, 80)).toBe(20)
      expect(calculateVariance(100, 120)).toBe(-20)
    })

    it('should return null for invalid inputs', () => {
      expect(calculateVariance(null, 80)).toBe(null)
      expect(calculateVariance(100, null)).toBe(null)
    })
  })

  describe('getVariancePercentage', () => {
    it('should calculate percentage variance correctly', () => {
      expect(getVariancePercentage(100, 120)).toBe(20)
      expect(getVariancePercentage(100, 80)).toBe(-20)
    })

    it('should return null for invalid inputs', () => {
      expect(getVariancePercentage(null, 80)).toBe(null)
      expect(getVariancePercentage(100, null)).toBe(null)
    })
  })

  describe('validateEmail', () => {
    it('should validate correct email addresses', () => {
      expect(validateEmail('test@example.com')).toBe(true)
      expect(validateEmail('user.name@domain.co.uk')).toBe(true)
    })

    it('should reject invalid email addresses', () => {
      expect(validateEmail('invalid-email')).toBe(false)
      expect(validateEmail('test@')).toBe(false)
      expect(validateEmail('@domain.com')).toBe(false)
    })
  })

  describe('validateRequired', () => {
    it('should validate required fields', () => {
      expect(validateRequired('value')).toBe(true)
      expect(validateRequired(0)).toBe(true)
      expect(validateRequired(false)).toBe(true)
    })

    it('should reject empty values', () => {
      expect(validateRequired(null)).toBe(false)
      expect(validateRequired(undefined)).toBe(false)
      expect(validateRequired('')).toBe(false)
    })
  })

  describe('validateNumber', () => {
    it('should validate numbers', () => {
      expect(validateNumber(123)).toBe(true)
      expect(validateNumber(0)).toBe(true)
      expect(validateNumber(-123)).toBe(true)
      expect(validateNumber(123.45)).toBe(true)
    })

    it('should reject non-numbers', () => {
      expect(validateNumber('abc')).toBe(false)
      expect(validateNumber(NaN)).toBe(false)
      expect(validateNumber(Infinity)).toBe(false)
    })
  })

  describe('validatePositiveNumber', () => {
    it('should validate positive numbers', () => {
      expect(validatePositiveNumber(123)).toBe(true)
      expect(validatePositiveNumber(0.1)).toBe(true)
    })

    it('should reject zero and negative numbers', () => {
      expect(validatePositiveNumber(0)).toBe(false)
      expect(validatePositiveNumber(-123)).toBe(false)
    })
  })

  describe('isEmpty', () => {
    it('should detect empty values', () => {
      expect(isEmpty(null)).toBe(true)
      expect(isEmpty(undefined)).toBe(true)
      expect(isEmpty('')).toBe(true)
      expect(isEmpty('   ')).toBe(true)
      expect(isEmpty([])).toBe(true)
      expect(isEmpty({})).toBe(true)
    })

    it('should detect non-empty values', () => {
      expect(isEmpty('value')).toBe(false)
      expect(isEmpty([1, 2, 3])).toBe(false)
      expect(isEmpty({ key: 'value' })).toBe(false)
      expect(isEmpty(0)).toBe(false)
    })
  })

  describe('capitalizeFirst', () => {
    it('should capitalize first letter', () => {
      expect(capitalizeFirst('hello')).toBe('Hello')
      expect(capitalizeFirst('HELLO')).toBe('HELLO')
    })

    it('should handle empty strings', () => {
      expect(capitalizeFirst('')).toBe('')
      expect(capitalizeFirst(null)).toBe('')
    })
  })

  describe('truncateText', () => {
    it('should truncate long text', () => {
      const longText = 'This is a very long text that should be truncated'
      expect(truncateText(longText, 20)).toBe('This is a very long ...')
    })

    it('should not truncate short text', () => {
      const shortText = 'Short text'
      expect(truncateText(shortText, 20)).toBe('Short text')
    })

    it('should handle null/undefined', () => {
      expect(truncateText(null)).toBe(null)
      expect(truncateText(undefined)).toBe(undefined)
    })
  })
})

