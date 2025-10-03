// Represents a single row of data from the Excel file.
// It's a flexible object where keys are column headers (string) and values can be anything.
export type DataRow = Record<string, any>;

// Defines the structure for sorting configuration.
export interface SortConfig {
  key: string;
  direction: 'ascending' | 'descending';
}

// Defines the structure for an employee's leave balance record.
export interface EmployeeRecord {
  id: string;
  name: string;
  balance: number;
  username: string;
  password: string;
  photo?: string; // To store base64 encoded profile picture
  priorHourlyBalance?: number; // Pre-existing hourly leave balance
  workdayHours: number; // Daily work hours for hourly leave calculation
}
