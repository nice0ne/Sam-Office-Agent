// Test setup configuration
import { beforeEach } from 'vitest';

beforeEach(() => {
  // Clear localStorage before each test
  if (typeof localStorage !== 'undefined') {
    localStorage.clear();
  }
});
