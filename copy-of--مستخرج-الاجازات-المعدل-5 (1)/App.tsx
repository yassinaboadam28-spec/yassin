import React, { useState, useMemo, useCallback, useRef, useEffect } from 'react';
import type { DataRow, SortConfig, EmployeeRecord } from './types';
import { AmiriFont } from './AmiriFont';
import { employeePhotos } from './employeePhotos';

// Make external libraries available, assuming they're loaded from a CDN.
declare const XLSX: any;
declare const jspdf: any;
declare const html2canvas: any;

// --- New types for structured summary ---
interface LeaveSummary {
  type: string;
  dayCount: number;
  hourCount: number;
  dateDetails: string;
}

interface EmployeeSummary {
  name: string;
  leaves: LeaveSummary[];
  initialBalance?: number;
  photo?: string; // Add photo to summary
}

// --- Type for Monthly Report ---
interface MonthlyReportRow {
    name: string;
    initialBalance: number;
    regularLeaves: { count: number; dates: string };
    hourlyLeaves: { days: number; hours: number };
    sickLeave: { dateRange: string };
    longLeave: { dateRange: string };
    finalBalance: number;
}

// --- Storage Keys ---
const EMPLOYEES_STORAGE_KEY = 'employeeLeaveBalances';
const LEAVE_DATA_STORAGE_KEY = 'leaveDataRecords';
const PROCESSED_FILES_STORAGE_KEY = 'processedFileNames';
const LAST_BALANCE_UPDATE_KEY = 'lastBalanceUpdateMarker';


// --- SVG Icon Components (defined outside the main component to prevent re-creation on re-renders) ---

const UploadIcon: React.FC<{ className?: string }> = ({ className }) => (
  <svg className={className} xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={1.5} stroke="currentColor">
    <path strokeLinecap="round" strokeLinejoin="round" d="M12 16.5V9.75m0 0l-3.75 3.75M12 9.75l3.75 3.75M3 17.25V21h18v-3.75M4.5 12.75l7.5-7.5 7.5 7.5" />
  </svg>
);

const DownloadIcon: React.FC<{ className?: string }> = ({ className }) => (
  <svg className={className} xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={1.5} stroke="currentColor">
    <path strokeLinecap="round" strokeLinejoin="round" d="M3 16.5v2.25A2.25 2.25 0 005.25 21h13.5A2.25 2.25 0 0021 18.75V16.5M16.5 12L12 16.5m0 0L7.5 12m4.5 4.5V3" />
  </svg>
);

const TrashIcon: React.FC<{ className?: string }> = ({ className }) => (
    <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={1.5} stroke="currentColor" className={className}>
        <path strokeLinecap="round" strokeLinejoin="round" d="M14.74 9l-.346 9m-4.788 0L9.26 9m9.968-3.21c.342.052.682.107 1.022.166m-1.022-.165L18.16 19.673a2.25 2.25 0 01-2.244 2.077H8.084a2.25 2.25 0 01-2.244-2.077L4.772 5.79m14.456 0a48.108 48.108 0 00-3.478-.397m-12 .562c.34-.059.68-.114 1.022-.165m0 0a48.11 48.11 0 013.478-.397m7.5 0v-.916c0-1.18-.91-2.144-2.09-2.201a51.964 51.964 0 00-3.32 0c-1.18.057-2.09 1.022-2.09 2.201v.916m7.5 0a48.667 48.667 0 00-7.5 0" />
    </svg>
);


const SearchIcon: React.FC<{ className?: string }> = ({ className }) => (
    <svg className={className} xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={1.5} stroke="currentColor">
        <path strokeLinecap="round" strokeLinejoin="round" d="M21 21l-5.197-5.197m0 0A7.5 7.5 0 105.196 5.196a7.5 7.5 0 0010.607 10.607z" />
    </svg>
);

const DocumentTextIcon: React.FC<{ className?: string }> = ({ className }) => (
    <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={1.5} stroke="currentColor" className={className}>
        <path strokeLinecap="round" strokeLinejoin="round" d="M19.5 14.25v-2.625a3.375 3.375 0 00-3.375-3.375h-1.5A1.125 1.125 0 0113.5 7.125v-1.5a3.375 3.375 0 00-3.375-3.375H8.25m0 12.75h7.5m-7.5 3H12M10.5 2.25H5.625c-.621 0-1.125.504-1.125 1.125v17.25c0 .621.504 1.125 1.125 1.125h12.75c.621 0 1.125-.504 1.125-1.125V11.25a9 9 0 00-9-9z" />
    </svg>
);

const TableCellsIcon: React.FC<{ className?: string }> = ({ className }) => (
    <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={1.5} stroke="currentColor" className={className}>
        <path strokeLinecap="round" strokeLinejoin="round" d="M3.375 19.5h17.25m-17.25 0a1.125 1.125 0 01-1.125-1.125v-1.5c0-.621.504-1.125 1.125-1.125h17.25c.621 0 1.125.504 1.125 1.125v1.5c0 .621-.504 1.125-1.125 1.125m-17.25 0h.008v.015h-.008v-.015zm1.5 0h.008v.015h-.008v-.015zm1.5 0h.008v.015h-.008v-.015zm1.5 0h.008v.015h-.008v-.015zm1.5 0h.008v.015h-.008v-.015zm1.5 0h.008v.015h-.008v-.015zm1.5 0h.008v.015h-.008v-.015zm1.5 0h.008v.015h-.008v-.015zm-16.5-3.375h17.25m-17.25 0a1.125 1.125 0 01-1.125-1.125v-1.5c0-.621.504-1.125 1.125-1.125h17.25c.621 0 1.125.504 1.125 1.125v1.5c0 .621-.504 1.125-1.125 1.125m-17.25 0h.008v.015h-.008v-.015zm1.5 0h.008v.015h-.008v-.015zm1.5 0h.008v.015h-.008v-.015zm1.5 0h.008v.015h-.008v-.015zm1.5 0h.008v.015h-.008v-.015zm1.5 0h.008v.015h-.008v-.015zm1.5 0h.008v.015h-.008v-.015zm1.5 0h.008v.015h-.008v-.015zm-16.5-3.375h17.25m-17.25 0A1.125 1.125 0 012.25 8.25v-1.5c0-.621.504-1.125 1.125-1.125h17.25c.621 0 1.125.504 1.125 1.125v1.5c0 .621-.504 1.125-1.125 1.125m-17.25 0h.008v.015h-.008V9.75zm1.5 0h.008v.015h-.008V9.75zm1.5 0h.008v.015h-.008V9.75zm1.5 0h.008v.015h-.008V9.75zm1.5 0h.008v.015h-.008V9.75zm1.5 0h.008v.015h-.008V9.75zm1.5 0h.008v.015h-.008V9.75zm1.5 0h.008v.015h-.008V9.75z" />
    </svg>
);

const ClipboardDocumentIcon: React.FC<{ className?: string }> = ({ className }) => (
    <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={1.5} stroke="currentColor" className={className}>
      <path strokeLinecap="round" strokeLinejoin="round" d="M15.666 3.888A2.25 2.25 0 0013.5 2.25h-3c-1.03 0-1.9.693-2.166 1.638m7.332 0c.055.194.084.4.084.612v0a2.25 2.25 0 01-2.25 2.25h-1.5a2.25 2.25 0 01-2.25-2.25v0c0-.212.03-.418.084-.612m7.332 0c.646.049 1.288.11 1.927.184 1.1.128 1.907 1.077 1.907 2.185V19.5a2.25 2.25 0 01-2.25 2.25H6.75A2.25 2.25 0 014.5 19.5V6.257c0-1.108.806-2.057 1.907-2.185a48.208 48.208 0 011.927-.184" />
    </svg>
);

const ClockIcon: React.FC<{ className?: string }> = ({ className }) => (
    <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={1.5} stroke="currentColor" className={className}>
        <path strokeLinecap="round" strokeLinejoin="round" d="M12 6v6h4.5m4.5 0a9 9 0 11-18 0 9 9 0 0118 0z" />
    </svg>
);

const BarsArrowUpIcon: React.FC<{ className?: string }> = ({ className }) => (
    <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={1.5} stroke="currentColor" className={className}>
        <path strokeLinecap="round" strokeLinejoin="round" d="M3 4.5h14.25M3 9h9.75M3 13.5h5.25m5.25-.75L17.25 9m0 0L21 12.75M17.25 9v12" />
    </svg>
);

const AlphabeticalSortIcon: React.FC<{ className?: string }> = ({ className }) => (
    <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={1.5} stroke="currentColor" className={className}>
      <path strokeLinecap="round" strokeLinejoin="round" d="M3.75 6.75h16.5M3.75 12h16.5m-16.5 5.25H12" />
    </svg>
);

const BriefcaseIcon: React.FC<{ className?: string }> = ({ className }) => (
    <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={1.5} stroke="currentColor" className={className}>
        <path strokeLinecap="round" strokeLinejoin="round" d="M20.25 14.15v4.075c0 1.313-.943 2.5-2.206 2.5H6.081c-1.262 0-2.206-1.187-2.206-2.5v-4.075m16.35 0c.225.045.45.08.675.112v-4.075c0-1.313-.943-2.5-2.206-2.5h-12.17c-1.263 0-2.206 1.187-2.206 2.5v4.075c.225-.032.45-.067.675-.112M16.5 7.5h-9" />
    </svg>
);

const HeartIcon: React.FC<{ className?: string }> = ({ className }) => (
    <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={1.5} stroke="currentColor" className={className}>
        <path strokeLinecap="round" strokeLinejoin="round" d="M21 8.25c0-2.485-2.099-4.5-4.688-4.5-1.935 0-3.597 1.126-4.312 2.733-.715-1.607-2.377-2.733-4.313-2.733C5.1 3.75 3 5.765 3 8.25c0 7.22 9 12 9 12s9-4.78 9-12z" />
    </svg>
);

const PrinterIcon: React.FC<{ className?: string }> = ({ className }) => (
    <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={1.5} stroke="currentColor" className={className}>
        <path strokeLinecap="round" strokeLinejoin="round" d="M6.72 13.829c-.24.03-.48.062-.72.096m.72-.096a42.415 42.415 0 0110.56 0m-10.56 0L6 3.369m0 0c.071.01.141.02.21.031m-1.215 10.518a42.597 42.597 0 01-2.073-.09m2.073.09a42.415 42.415 0 0010.56 0m-10.56 0c.316.05.632.095.948.135M16.5 10.5V6.75a4.5 4.5 0 10-9 0v3.75m-.75 11.25h10.5a2.25 2.25 0 002.25-2.25v-6.75a2.25 2.25 0 00-2.25-2.25H6.75a2.25 2.25 0 00-2.25 2.25v6.75a2.25 2.25 0 002.25 2.25z" />
    </svg>
);

const PdfIcon: React.FC<{ className?: string }> = ({ className }) => (
  <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={1.5} stroke="currentColor" className={className}>
    <path strokeLinecap="round" strokeLinejoin="round" d="M19.5 14.25v-2.625a3.375 3.375 0 00-3.375-3.375h-1.5A1.125 1.125 0 0113.5 7.125v-1.5a3.375 3.375 0 00-3.375-3.375H8.25m.158 10.302L9 18.25m1.5-3.948L10.5 18.25m0 0l.158.25m-.158-.25L9.342 16.5m.208 1.75h.208m-3.375 0h3.375m-3.375 0h.008v.015h-.008v-.015z" />
    <path strokeLinecap="round" strokeLinejoin="round" d="M19.5 14.25v-2.625a3.375 3.375 0 00-3.375-3.375h-1.5A1.125 1.125 0 0113.5 7.125v-1.5a3.375 3.375 0 00-3.375-3.375H8.25m2.25 0H5.625c-.621 0-1.125.504-1.125 1.125v17.25c0 .621.504 1.125 1.125 1.125h12.75c.621 0 1.125-.504 1.125-1.125V11.25a9 9 0 00-9-9z" />
  </svg>
);

// --- New Icons for Summary Cards ---
const UserGroupIcon: React.FC<{ className?: string }> = ({ className }) => (
    <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={1.5} stroke="currentColor" className={className}>
        <path strokeLinecap="round" strokeLinejoin="round" d="M18 18.72a9.094 9.094 0 00-12 0m12 0a9.094 9.094 0 00-12 0m12 0v.006M18 18.72v.006m-12 0v.006m0-10.74a6 6 0 0112 0v1.5a6 6 0 10-12 0v-1.5zm12 0a6 6 0 00-12 0v1.5a6 6 0 1012 0v-1.5z" />
    </svg>
);

const UserCircleIcon: React.FC<{ className?: string }> = ({ className }) => (
    <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={1.5} stroke="currentColor" className={className}>
      <path strokeLinecap="round" strokeLinejoin="round" d="M17.982 18.725A7.488 7.488 0 0012 15.75a7.488 7.488 0 00-5.982 2.975m11.963 0a9 9 0 10-11.963 0m11.963 0A8.966 8.966 0 0112 21a8.966 8.966 0 01-5.982-2.275M15 9.75a3 3 0 11-6 0 3 3 0 016 0z" />
    </svg>
);


const UsersIcon: React.FC<{ className?: string }> = ({ className }) => (
    <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={1.5} stroke="currentColor" className={className}>
        <path strokeLinecap="round" strokeLinejoin="round" d="M15 19.128a9.38 9.38 0 002.625.372 9.337 9.337 0 004.121-.952 4.125 4.125 0 00-7.533-2.493M15 19.128v-.003c0-1.113-.285-2.16-.786-3.07M15 19.128v.106A12.318 12.318 0 018.624 21c-2.331 0-4.512-.645-6.374-1.766l-.001-.109a6.375 6.375 0 0111.964-4.663M12 3.375c-3.418 0-6.167 2.023-6.167 4.5s2.75 4.5 6.167 4.5 6.167-2.023 6.167-4.5S15.418 3.375 12 3.375z" />
    </svg>
);

const PencilIcon: React.FC<{ className?: string }> = ({ className }) => (
    <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={1.5} stroke="currentColor" className={className}>
        <path strokeLinecap="round" strokeLinejoin="round" d="M16.862 4.487l1.687-1.688a1.875 1.875 0 112.652 2.652L10.582 16.07a4.5 4.5 0 01-1.897 1.13L6 18l.8-2.685a4.5 4.5 0 011.13-1.897l8.932-8.931zm0 0L19.5 7.125M18 14v4.75A2.25 2.25 0 0115.75 21H5.25A2.25 2.25 0 013 18.75V8.25A2.25 2.25 0 015.25 6H10" />
    </svg>
);

const CalendarDaysIcon: React.FC<{ className?: string }> = ({ className }) => (
    <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={1.5} stroke="currentColor" className={className}>
      <path strokeLinecap="round" strokeLinejoin="round" d="M6.75 3v2.25M17.25 3v2.25M3 18.75V7.5a2.25 2.25 0 012.25-2.25h13.5A2.25 2.25 0 0121 7.5v11.25m-18 0A2.25 2.25 0 005.25 21h13.5A2.25 2.25 0 0021 18.75m-18 0h18" />
    </svg>
);

const TrophyIcon: React.FC<{ className?: string }> = ({ className }) => (
    <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={1.5} stroke="currentColor" className={className}>
      <path strokeLinecap="round" strokeLinejoin="round" d="M16.5 18.75h-9a9 9 0 119 0zM16.5 18.75a9 9 0 10-9 0m9 0h-9m9 0h-9M9 13.5v6.375m6-6.375v6.375m-3-3.375V18.75m0-12.75a3 3 0 00-3-3H9a3 3 0 00-3 3v.75m6 .75v-3.75m-6 3.75v-3.75m3 3.75v-3.75M9 3h6" />
    </svg>
);

const ArrowPathIcon: React.FC<{ className?: string }> = ({ className }) => (
    <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={1.5} stroke="currentColor" className={className}>
        <path strokeLinecap="round" strokeLinejoin="round" d="M16.023 9.348h4.992v-.001M2.985 19.644v-4.992m0 0h4.992m-4.993 0l3.181 3.183a8.25 8.25 0 0011.664 0l3.181-3.183m-4.991-2.696a8.25 8.25 0 00-11.664 0l-3.181 3.183" />
    </svg>
);

const MinusCircleIcon: React.FC<{ className?: string }> = ({ className }) => (
    <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={1.5} stroke="currentColor" className={className}>
        <path strokeLinecap="round" strokeLinejoin="round" d="M15 12H9m12 0a9 9 0 11-18 0 9 9 0 0118 0z" />
    </svg>
);


const SortIcon: React.FC<{ direction?: 'ascending' | 'descending' }> = ({ direction }) => {
  if (!direction) {
    return (
      <svg className="w-4 h-4 text-gray-400" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={2} stroke="currentColor">
        <path strokeLinecap="round" strokeLinejoin="round" d="M8.25 15L12 18.75 15.75 15m-7.5-6L12 5.25 15.75 9" />
      </svg>
    );
  }
  if (direction === 'ascending') {
    return (
      <svg className="w-4 h-4" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={2} stroke="currentColor">
        <path strokeLinecap="round" strokeLinejoin="round" d="M4.5 15.75l7.5-7.5 7.5 7.5" />
      </svg>
    );
  }
  return (
    <svg className="w-4 h-4" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={2} stroke="currentColor">
      <path strokeLinecap="round" strokeLinejoin="round" d="M19.5 8.25l-7.5 7.5-7.5-7.5" />
    </svg>
  );
};

/**
 * Converts Western/ASCII numerals in a string to Eastern Arabic numerals.
 */
const toArabicNumerals = (input: string | number): string => {
    const strInput = String(input);
    const arabicNumerals = ['٠', '١', '٢', '٣', '٤', '٥', '٦', '٧', '٨', '٩'];
    return strInput.replace(/[0-9]/g, (digit) => arabicNumerals[parseInt(digit)]);
};

/**
 * Generates a random alphanumeric password.
 */
const generateRandomPassword = (length: number = 7): string => {
  const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
  let result = '';
  for (let i = 0; i < length; i++) {
    result += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return result;
};

/**
 * Generates a random two-digit username.
 */
const generateRandomUsername = (): string => {
  return String(Math.floor(Math.random() * 100)).padStart(2, '0');
};

/**
 * Resizes an image file and converts it to a base64 encoded JPEG data URL.
 * This helps keep the storage size manageable.
 */
const processImageFile = (file: File, maxWidth: number = 128, maxHeight: number = 128): Promise<string> => {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.readAsDataURL(file);
        reader.onload = (event) => {
            if (!event.target?.result) {
                return reject(new Error("Failed to read file."));
            }
            const img = new Image();
            img.src = event.target.result as string;
            img.onload = () => {
                const canvas = document.createElement('canvas');
                let { width, height } = img;

                // Calculate the new dimensions
                if (width > height) {
                    if (width > maxWidth) {
                        height = Math.round((height * maxWidth) / width);
                        width = maxWidth;
                    }
                } else {
                    if (height > maxHeight) {
                        width = Math.round((width * maxHeight) / height);
                        height = maxHeight;
                    }
                }

                canvas.width = width;
                canvas.height = height;
                const ctx = canvas.getContext('2d');
                if (!ctx) {
                    return reject(new Error('Could not get canvas context'));
                }
                ctx.drawImage(img, 0, 0, width, height);
                
                // Convert to data URL with JPEG compression
                resolve(canvas.toDataURL('image/jpeg', 0.85));
            };
            img.onerror = (error) => reject(error);
        };
        reader.onerror = (error) => reject(error);
    });
};


/**
 * Formats a number of days into a grammatically correct Arabic string.
 */
const formatDaysArabic = (count: number): string => {
    const num = Number(count);
    if (isNaN(num) || num <= 0) return '٠ يوم';
    if (num === 1) return 'يوم واحد';
    if (num === 2) return 'يومان';
    if (num >= 3 && num <= 10) {
        const words = ['ثلاثة', 'أربعة', 'خمسة', 'ستة', 'سبعة', 'ثمانية', 'تسعة', 'عشرة'];
        return `${words[num - 3]} أيام`;
    }
    // For numbers > 10, the format is [number] + singular noun (accusative)
    return `${toArabicNumerals(num)} يومًا`;
};

/**
 * Formats leave counts (days and hours) into an Arabic string with Arabic numerals.
 */
const formatLeaveCount = (days: number, hours: number): string => {
    const parts: string[] = [];
    if (days > 0) parts.push(`${toArabicNumerals(days)} يوم`);
    const formattedHours = parseFloat(hours.toFixed(2));
    if (formattedHours > 0) parts.push(`${toArabicNumerals(formattedHours)} ساعة`);
    return parts.length > 0 ? parts.join(' و ') : 'لا يوجد';
};


/**
 * Processes raw data from SheetJS to handle messy and unstructured Excel files.
 * It identifies columns by content, not just headers, and fills in missing employee names
 * for consecutive rows.
 */
const processAndCleanData = (rawData: DataRow[]): { cleanData: DataRow[], cleanHeaders: string[] } => {
    if (rawData.length < 1) {
        return { cleanData: [], cleanHeaders: [] };
    }

    const daysOfWeek = ['الاحد', 'الاثنين', 'الثلاثاء', 'الاربعاء', 'الخميس', 'الجمعة', 'السبت'];
    const leaveTypesKeywords = ['اجازة', 'زمنية', 'رصد'];
    const headers = Object.keys(rawData[0] || {});

    // Step 1: Score each column based on content matching
    const headerScores: Record<string, Record<string, number>> = {};
    headers.forEach(h => {
        headerScores[h] = { name: 0, date: 0, day: 0, type: 0, value: 0, id: 0, other: 0 };
    });

    const sampleSize = Math.min(rawData.length, 50); // Use a larger sample size for better accuracy

    for (let i = 0; i < sampleSize; i++) {
        const row = rawData[i];
        for (const header of headers) {
            const val = row[header];
            if (val == null || String(val).trim() === '') continue;

            const valueStr = String(val).trim();

            if (val instanceof Date || valueStr.match(/^\d{1,2}[-\/]\d{1,2}[-\/]\d{2,4}$/)) {
                headerScores[header].date++;
            } else if (daysOfWeek.includes(valueStr)) {
                headerScores[header].day++;
            } else if (leaveTypesKeywords.some(keyword => valueStr.toLowerCase().includes(keyword))) {
                headerScores[header].type++;
            } else if (!isNaN(parseFloat(valueStr)) && isFinite(Number(val))) {
                const num = Number(val);
                
                if (valueStr.includes('.')) {
                    headerScores[header].value += 3; 
                } else { 
                    if (num > 0 && num <= 24) {
                        headerScores[header].value++;
                    }
                    
                    if (num > 24) {
                        headerScores[header].id += 2;
                    } else if (num >= 1) {
                        headerScores[header].id++;
                    }
                }
            } else if (valueStr.includes(' ') && isNaN(Number(valueStr)) && valueStr.length > 5) {
                headerScores[header].name++;
            } else {
                headerScores[header].other++;
            }
        }
    }
    
    // Step 2: Find the best matching header for each data type, preventing duplicates
    const findBestHeader = (type: 'name' | 'date' | 'day' | 'type' | 'value', usedHeaders: Set<string>): string | null => {
        let bestHeader: string | null = null;
        let maxScore = 0;

        for (const header of headers) {
            if (usedHeaders.has(header)) continue;

            if (headerScores[header][type] > maxScore) {
                maxScore = headerScores[header][type];
                bestHeader = header;
            }
        }
        
        return maxScore > 0 ? bestHeader : null;
    };

    const usedHeaders = new Set<string>();
    const headerMapping: Record<string, string | null> = {};
    const typesToFind: ('date' | 'type' | 'name' | 'day' | 'value')[] = ['date', 'type', 'name', 'day', 'value'];

    typesToFind.forEach(type => {
        const header = findBestHeader(type, usedHeaders);
        headerMapping[type] = header;
        if (header) {
            usedHeaders.add(header);
        }
    });

    const { name: nameHeader, date: dateHeader, type: typeHeader, day: dayHeader, value: valueHeader } = headerMapping;
    
    // Step 3: Validate that required columns were found
    if (!nameHeader || !dateHeader || !typeHeader) {
        const missing = [];
        if (!nameHeader) missing.push("'الاسم'");
        if (!dateHeader) missing.push("'التاريخ'");
        if (!typeHeader) missing.push("'نوع الاجازة'");
        throw new Error(`لم نتمكن من تحديد الأعمدة التالية تلقائيًا: ${missing.join('، ')}. يرجى التأكد من أن الملف يحتوي على هذه الأعمدة.`);
    }

    // Step 4: Process and clean the data row by row
    const cleanData: DataRow[] = [];
    let currentName: string = '';

    rawData.forEach(row => {
        const name = (row[nameHeader!] ?? '').toString().trim();
        const date = row[dateHeader!];
        const type = (row[typeHeader!] ?? '').toString().trim();

        if (name) {
            currentName = name;
        }

        if (date && type) {
            let formattedDate: string;
            if (date instanceof Date) {
                const d = String(date.getDate()).padStart(2, '0');
                const m = String(date.getMonth() + 1).padStart(2, '0');
                const y = date.getFullYear();
                formattedDate = `${d}-${m}-${y}`;
            } else {
                // Handle Excel's numeric date format
                if (typeof date === 'number' && date > 1) {
                    const excelEpoch = new Date(1899, 11, 30);
                    const jsDate = new Date(excelEpoch.getTime() + date * 86400000);
                    const d = String(jsDate.getUTCDate()).padStart(2, '0');
                    const m = String(jsDate.getUTCMonth() + 1).padStart(2, '0');
                    const y = jsDate.getUTCFullYear();
                    formattedDate = `${d}-${m}-${y}`;
                } else {
                     formattedDate = String(date);
                }
            }

            const newRow: DataRow = {
                'الاسم': currentName,
                'التاريخ': formattedDate,
                'يوم العمل': dayHeader && row[dayHeader] ? row[dayHeader] : '',
                'نوع الاجازة': type,
                'القيمة': valueHeader && row[valueHeader] != null ? row[valueHeader] : '',
            };
            cleanData.push(newRow);
        }
    });

    const cleanHeaders = ['الاسم', 'التاريخ', 'يوم العمل', 'نوع الاجازة', 'القيمة'];
    return { cleanData, cleanHeaders };
};


// --- Main Application Component ---

export const App: React.FC = () => {
  const [data, setData] = useState<DataRow[]>([]);
  const [processedFiles, setProcessedFiles] = useState<string[]>([]);
  const [headers, setHeaders] = useState<string[]>(['الاسم', 'التاريخ', 'يوم العمل', 'نوع الاجازة', 'القيمة']);
  const [searchTerm, setSearchTerm] = useState('');
  const [sortConfig, setSortConfig] = useState<SortConfig | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [isLoading, setIsLoading] = useState<boolean>(false);
  const [view, setView] = useState<'table' | 'summary' | 'rankedSummary' | 'employeeManagement' | 'upload' | 'monthlyReport'>('upload');
  const [summary, setSummary] = useState<EmployeeSummary[]>([]);
  const [showCopyNotification, setShowCopyNotification] = useState<boolean>(false);
  const [notification, setNotification] = useState<string | null>(null);
  const [rankedSortOrder, setRankedSortOrder] = useState<'byDays' | 'alphabetical'>('byDays');
  
  // --- New State for Summary Filtering ---
  const [summarySearchTerm, setSummarySearchTerm] = useState('');
  const [selectedYear, setSelectedYear] = useState('all');
  const [selectedMonth, setSelectedMonth] = useState('all');
  
  // --- New State for Monthly Report ---
  const [reportYear, setReportYear] = useState(new Date().getFullYear());
  const [reportMonth, setReportMonth] = useState(new Date().getMonth() + 1);
  const [monthlyReportData, setMonthlyReportData] = useState<MonthlyReportRow[]>([]);


  // --- Authentication State ---
  const [currentUser, setCurrentUser] = useState<EmployeeRecord | 'admin' | null>(null);
  const [usernameInput, setUsernameInput] = useState<string>('');
  const [passwordInput, setPasswordInput] = useState<string>('');
  const [loginError, setLoginError] = useState<string | null>(null);
  const ADMIN_USERNAME = 'admin';
  const ADMIN_PASSWORD = '2781';

  // --- Employee Management State ---
  const [employees, setEmployees] = useState<EmployeeRecord[]>([]);
  const [editingEmployee, setEditingEmployee] = useState<EmployeeRecord | null>(null);
  const [editingEmployeePhotoFile, setEditingEmployeePhotoFile] = useState<File | null>(null);
  const [newEmployeeName, setNewEmployeeName] = useState('');
  const [newEmployeeBalance, setNewEmployeeBalance] = useState('');
  const [newEmployeeWorkdayHours, setNewEmployeeWorkdayHours] = useState('7');
  const [newEmployeeUsername, setNewEmployeeUsername] = useState(generateRandomUsername());
  const [newEmployeePassword, setNewEmployeePassword] = useState(generateRandomPassword());
  const [newEmployeePhoto, setNewEmployeePhoto] = useState<string | null>(null);
  const [newEmployeePhotoFile, setNewEmployeePhotoFile] = useState<File | null>(null);
  const [employeeSearchTerm, setEmployeeSearchTerm] = useState('');
  const [isImageProcessing, setIsImageProcessing] = useState(false);


  const fileInputRef = useRef<HTMLInputElement>(null);
  const summaryContentRef = useRef<HTMLDivElement>(null);
  
    const initialEmployeeNames = [
    'اثيب عبد الزهرة مجيد', 'احمد امير حسين', 'احمد سمير محمد', 'احمد ناجح رزاق', 'ازهر ثامر ربح', 'استبرق جابر حميد', 'امير خالد هادي', 'انغام عبد الزهرة مجيد', 'انوار فاضل طراد', 'باسم عباس حسين', 'باسم علي محمد', 'باقر عادل احمد', 'تقى عبد الحسن حمودي', 'جاسم حازم محمد', 'جلال حسن هادي', 'حسن كريم باجي', 'حسنين عبد الرزاق', 'حسين علي عبد الأمير', 'حكمت شوكت عبد الحمزة', 'حمزة سعد حمودي', 'حيدر عباس حسن', 'حيدر كاظم محمد سعيد', 'خليل كريم عباس', 'رائد كامل عبد اليمة', 'راضي حمودي سلطان', 'رغد عبد الزهرة مجيد', 'سجاد محمد علي طالب', 'سعد حمودي سلطان', 'سعيد عبد الحسن حمودي', 'سهاد جابر محمد', 'ضرغام جهادي خضير', 'عباس سعد حمودي', 'عباس كاظم عبد الاخوه', 'عبد الرضا نجم عبد', 'عقيل مسلم عبد المحسن', 'علي باسم حسن', 'علي رسول محي', 'علي عبد الكريم حميد', 'علي كاظم قربون', 'علي محمد باجي', 'عمار ناصر محمد', 'غسان نزار ضياء', 'قاسم محمد رضا مجيد', 'كرار عدي عبد الحسين', 'كواكب عزيز حمزة', 'كوثر هادي عطيه', 'ليث محمد علي حمودي', 'مجيد محمد رضا مجيد', 'محمد جواد كاظم عباس', 'محمد راضي حمودي', 'محمد رضا عبود', 'محمد صلاح محمد', 'محمد عبد الرضا تركي', 'محمد علاء محمد', 'محمد فاضل شاكر', 'محمد محمد رضا مجيد', 'مرتضى كريم موسى', 'مروة محمد علي', 'مسلم إبراهيم حسن', 'مصطفى راضي حمودي', 'مصطفى علي محمد', 'مصطفى محسن يعقوب', 'مها سعد حمودي', 'مهدي صالح هادي', 'مهند محمد رشاد جعفر', 'ناجح كاظم قربون', 'هدى عبد الحسن حمودي', 'هناء علي عبد الرضا', 'ياسر فائز جاسم', 'ياسر ياسين ناجي', 'ياسين رياض احمد', 'زهراء طه تقي', 'يحيى فارس محمد'
  ];

  const initialEmployeeBalances = [
    16, 26, 38, 61, 36, 4, 14, 20, 9, 41, 33, 27, 22, 13, 54, 44, 67, 17, 26, 27, 48, 67, 51, 21, 39, 15, 10, 27, 51, 4, 19, 18, 9, 24, 6, 5, 25, 21, 30, 24, -1, 21, 6, 0, 13, -13, 27, 16, 53, 35, 51, 54, 19, 67, 5, 22, 1, 10, 26, -1, 16, 14, 11, 14, 6, 56, 28, 18, 37, -9, -6, 12
  ];
  
  const initialHourlyBalances = [
    0, 2, 2, 0, 0, 3, 3, 5, 5, 6, 3, 6, 2, 0, 0, 0, 0, 3, 3, 0, 1, 0, 0, 0, 1, 0, 0, 0, 0, 4, 0, 2, 5, 0, 0, 3, 4, 0, 2, 1, 4, 0, 4, 5, 2, 3, 0, 5, 2, 4, 0, 0, 1, 0, 2, 6, 3, 3, 5, 1, 5, 5, 0, 4, 3, 0, 6, 0, 0, 1, 2, 0, 5
  ];

  // --- Data Persistence ---

  // Load all data from localStorage on initial mount
  useEffect(() => {
    try {
      // Employees
      let storedEmployees = localStorage.getItem(EMPLOYEES_STORAGE_KEY);
      if (!storedEmployees || JSON.parse(storedEmployees).length === 0) {
        const initialEmployees: EmployeeRecord[] = initialEmployeeNames.map((name, index) => {
           // Generate deterministic username and password based on employee's name and index
           const username = String(100 + index);
           const firstName = name.split(' ')[0];
           const password = `${firstName}${10 + index}`;

           return {
             id: `${Date.now()}-${index}`,
             name,
             balance: initialEmployeeBalances[index] ?? 0,
             username: username,
             password: password,
             photo: employeePhotos[index] ? `data:image/jpeg;base64,${employeePhotos[index]}` : undefined,
             priorHourlyBalance: initialHourlyBalances[index] ?? 0,
             workdayHours: 7, // Default workday hours
           };
        }).sort((a,b) => a.name.localeCompare(b.name, 'ar'));
        
        storedEmployees = JSON.stringify(initialEmployees);
        localStorage.setItem(EMPLOYEES_STORAGE_KEY, storedEmployees);
      }
      setEmployees(JSON.parse(storedEmployees));

      // Leave Data
      const storedData = localStorage.getItem(LEAVE_DATA_STORAGE_KEY);
      const parsedData = storedData ? JSON.parse(storedData) : [];
      setData(parsedData);

      // Processed Files
      const storedFiles = localStorage.getItem(PROCESSED_FILES_STORAGE_KEY);
      setProcessedFiles(storedFiles ? JSON.parse(storedFiles) : []);
      
      // Set initial view based on loaded data
      if (parsedData.length > 0) {
        setView('table');
      }

    } catch (err) {
      console.error("Failed to load data from localStorage", err);
      setError("فشل تحميل البيانات المحفوظة.");
    }
  }, []);

  // Save data to localStorage whenever it changes
  useEffect(() => {
    try {
      localStorage.setItem(EMPLOYEES_STORAGE_KEY, JSON.stringify(employees));
    } catch (err) {
      console.error("Failed to save employees to localStorage", err);
       if (err instanceof DOMException && (err.name === 'QuotaExceededError' || err.name === 'NS_ERROR_DOM_QUOTA_REACHED')) {
           setError("فشل حفظ البيانات. قد تكون مساحة التخزين ممتلئة. حاول استخدام صورة بحجم أصغر.");
      } else {
           setError("حدث خطأ غير متوقع أثناء حفظ بيانات الموظفين.");
      }
    }
  }, [employees]);
  
  useEffect(() => {
    try {
      localStorage.setItem(LEAVE_DATA_STORAGE_KEY, JSON.stringify(data));
    } catch (err) {
      console.error("Failed to save leave data to localStorage", err);
    }
  }, [data]);

  useEffect(() => {
    try {
      localStorage.setItem(PROCESSED_FILES_STORAGE_KEY, JSON.stringify(processedFiles));
    } catch (err) {
      console.error("Failed to save processed files list to localStorage", err);
    }
  }, [processedFiles]);
  
  const parseDate = useCallback((dateStr: string): Date => {
      const parts = dateStr.split(/[-/]/);
      if (parts.length === 3) {
          const [day, month, year] = parts.map(p => parseInt(p, 10));
          if (!isNaN(day) && !isNaN(month) && !isNaN(year)) {
              if (year > 1000) {
                  return new Date(Date.UTC(year, month - 1, day));
              }
              const fullYear = year > 50 ? 1900 + year : 2000 + year;
              return new Date(Date.UTC(fullYear, month - 1, day));
          }
      }
      return new Date(0);
  }, []);

  const generateSummary = useCallback((summaryData: DataRow[], employeeRecords: EmployeeRecord[]): EmployeeSummary[] => {
    const normalizeName = (name: string): string => {
        if (!name) return '';
        return name
            .trim()
            .replace(/\s+/g, '') // Remove all whitespace
            .replace(/[أإآ]/g, 'ا')  // Normalize Alef
            .replace(/ة/g, 'ه')    // Normalize Teh Marbuta
            .replace(/ى/g, 'ي');   // Normalize Yaa
    };

    const normalizedEmployeeMap = new Map<string, EmployeeRecord>();
    employeeRecords.forEach(emp => {
        normalizedEmployeeMap.set(normalizeName(emp.name), emp);
    });

    const nameResolutionCache = new Map<string, string>();

    const resolveCanonicalName = (sheetName: string): string => {
        if (!sheetName) return sheetName;
        if (nameResolutionCache.has(sheetName)) {
            return nameResolutionCache.get(sheetName)!;
        }

        const normalizedSheetName = normalizeName(sheetName);

        if (normalizedEmployeeMap.has(normalizedSheetName)) {
            const canonicalName = normalizedEmployeeMap.get(normalizedSheetName)!.name;
            nameResolutionCache.set(sheetName, canonicalName);
            return canonicalName;
        }

        const potentialMatches = employeeRecords.filter(emp => {
            const normalizedEmpName = normalizeName(emp.name);
            return normalizedEmpName.includes(normalizedSheetName) || normalizedSheetName.includes(normalizedEmpName);
        });

        if (potentialMatches.length === 1) {
            const canonicalName = potentialMatches[0].name;
            nameResolutionCache.set(sheetName, canonicalName);
            return canonicalName;
        } else if (potentialMatches.length > 1) {
            potentialMatches.sort((a, b) =>
                Math.abs(normalizeName(a.name).length - normalizedSheetName.length) -
                Math.abs(normalizeName(b.name).length - normalizedSheetName.length)
            );
            const canonicalName = potentialMatches[0].name;
            nameResolutionCache.set(sheetName, canonicalName);
            return canonicalName;
        }

        nameResolutionCache.set(sheetName, sheetName);
        return sheetName;
    };

    const groupedData: Record<string, Record<string, { date: string, value: number }[]>> = {};

    summaryData.forEach(row => {
        const nameFromSheet = row['الاسم'];
        if (!nameFromSheet) return;

        const canonicalName = resolveCanonicalName(nameFromSheet);
        const date = row['التاريخ'];
        let type = (row['نوع الاجازة'] || '').toString().trim();
        const value = parseFloat(String(row['القيمة']).trim()) || 0;

        if (!canonicalName || !date || !type) {
            return;
        }

        if (type.startsWith('زمنية')) {
            type = 'اجازة زمنية';
        } else {
            type = type.replace(/\s+/g, ' ').trim();
        }

        if (type.includes('رصد') && type.includes('مسائي')) {
            return;
        }

        if (!groupedData[canonicalName]) groupedData[canonicalName] = {};
        if (!groupedData[canonicalName][type]) groupedData[canonicalName][type] = [];

        groupedData[canonicalName][type].push({ date: String(date), value });
    });

    const allNames = new Set([...Object.keys(groupedData), ...employeeRecords.map(e => e.name)]);
    const sortedNames = Array.from(allNames).sort((a, b) => a.localeCompare(b, 'ar'));
    
    const finalSummary: EmployeeSummary[] = [];

    for (const name of sortedNames) {
        const employeeRecord = employeeRecords.find(emp => emp.name === name);
        const employeeSummary: EmployeeSummary = {
            name,
            leaves: [],
            initialBalance: employeeRecord?.balance,
            photo: employeeRecord?.photo
        };
        const employeeVacations = groupedData[name] || {};
        
        const priorHourlyBalance = employeeRecord?.priorHourlyBalance || 0;
        
        let workdayHours = employeeRecord?.workdayHours || 7; // Start with stored value or default
        // Dynamically check data from the sheet to override stored value
        const regularLeavesForHoursCheck = employeeVacations['اجازة اعتيادية'] || [];
        if (regularLeavesForHoursCheck.length > 0) {
            const typicalHours = regularLeavesForHoursCheck[0].value;
            if (typicalHours === 6 || typicalHours === 7) {
                workdayHours = typicalHours; // Override with actual data
            }
        }
        
        const leaveTypesToProcess = new Set(Object.keys(employeeVacations));
        if (priorHourlyBalance > 0) {
            leaveTypesToProcess.add('اجازة زمنية');
        }

        const sortedTypes = Array.from(leaveTypesToProcess).sort((a, b) => {
            const order = ['اجازة اعتيادية', 'اجازة مرضية', 'ملخص الزمنيات'];
            const indexA = order.indexOf(a);
            const indexB = order.indexOf(b);
            if(indexA > -1 && indexB > -1) return indexA - indexB;
            if(indexA > -1) return -1;
            if(indexB > -1) return 1;
            return a.localeCompare(b, 'ar');
        });

        for (const type of sortedTypes) {
            const entries = employeeVacations[type] || [];

            if (type === 'اجازة زمنية') {
                const totalHoursFromSheet = entries.reduce((sum, entry) => sum + entry.value, 0);
                const totalHours = totalHoursFromSheet + priorHourlyBalance;

                if (totalHours > 0) {
                    const days = Math.floor(totalHours / workdayHours);
                    const remainingHours = totalHours % workdayHours;
                    const formattedHours = parseFloat(remainingHours.toFixed(2));
                    
                    let dateDetails = '';
                    if (entries.length > 0) {
                        dateDetails = `إجمالي ${toArabicNumerals(totalHoursFromSheet)} ساعة عبر ${toArabicNumerals(entries.length)} إدخال`;
                    }
                    if (priorHourlyBalance > 0) {
                         if (dateDetails) dateDetails += ' + ';
                         dateDetails += `${toArabicNumerals(priorHourlyBalance)} ساعة رصيد سابق`;
                    }

                    employeeSummary.leaves.push({
                        type: 'ملخص الزمنيات',
                        dayCount: days,
                        hourCount: formattedHours,
                        dateDetails: dateDetails
                    });
                }
            } else if (type.includes('مرضية') || type.includes('طويلة')) {
                const sortedDates = entries
                    .map(e => parseDate(String(e.date)))
                    .filter(d => !isNaN(d.getTime()))
                    .sort((a, b) => a.getTime() - b.getTime());

                if (sortedDates.length > 0) {
                    const periods: { start: Date, end: Date }[] = [];
                    let currentPeriod = { start: sortedDates[0], end: sortedDates[0] };
                    
                    for (let i = 1; i < sortedDates.length; i++) {
                        const currentDate = sortedDates[i];
                        const prevDate = currentPeriod.end;
                        const diffDays = (currentDate.getTime() - prevDate.getTime()) / (1000 * 60 * 60 * 24);

                        if (diffDays === 1) {
                            currentPeriod.end = currentDate;
                        } else {
                            periods.push(currentPeriod);
                            currentPeriod = { start: currentDate, end: currentDate };
                        }
                    }
                    periods.push(currentPeriod);
                    
                    periods.forEach(period => {
                        const dayCount = Math.round((period.end.getTime() - period.start.getTime()) / (1000 * 60 * 60 * 24)) + 1;

                        const formatDate = (date: Date) => {
                            const d = String(date.getUTCDate()).padStart(2, '0');
                            const m = String(date.getUTCMonth() + 1).padStart(2, '0');
                            const y = date.getUTCFullYear();
                            return `${d}-${m}-${y}`;
                        };

                        let dateDetailsStr = '';
                        if (dayCount <= 1) {
                            dateDetailsStr = formatDate(period.start);
                        } else {
                            dateDetailsStr = `من ${formatDate(period.start)} إلى ${formatDate(period.end)}`;
                        }

                        employeeSummary.leaves.push({
                            type,
                            dayCount: dayCount,
                            hourCount: 0,
                            dateDetails: dateDetailsStr
                        });
                    });
                }
            } else {
                const count = entries.length;
                const sortedDates = entries
                    .map(e => parseDate(String(e.date)))
                    .sort((a, b) => a.getTime() - b.getTime());
                
                let dateDetailsStr = '';
                
                const datesByMonthYear: Record<string, number[]> = {};

                sortedDates.forEach(date => {
                    if (!isNaN(date.getTime())) {
                        const month = date.getUTCMonth() + 1;
                        const year = date.getUTCFullYear();
                        const key = `${year}-${String(month).padStart(2, '0')}`;

                        if (!datesByMonthYear[key]) {
                            datesByMonthYear[key] = [];
                        }
                        datesByMonthYear[key].push(date.getUTCDate());
                    }
                });

                const formattedDateParts = Object.keys(datesByMonthYear).map(key => {
                    const [year, month] = key.split('-');
                    const days = datesByMonthYear[key];
                    days.sort((a, b) => a - b);
                    const arabicDays = days.map(d => toArabicNumerals(d)).join('،');
                    const arabicMonth = toArabicNumerals(parseInt(month, 10));
                    const arabicYear = toArabicNumerals(year);
                    return `${arabicDays}/${arabicMonth}/${arabicYear}`;
                });
                
                if (formattedDateParts.length > 0) {
                     dateDetailsStr = formattedDateParts.join(' | ');
                }
                
                employeeSummary.leaves.push({
                    type,
                    dayCount: count,
                    hourCount: 0,
                    dateDetails: dateDetailsStr
                });
            }
        }
        finalSummary.push(employeeSummary);
    }

    return finalSummary;
  }, [parseDate]);

  // Regenerate summary whenever data or employees change
  useEffect(() => {
    if (data.length > 0 || employees.length > 0) {
        const newSummary = generateSummary(data, employees);
        setSummary(newSummary);
    }
  }, [data, employees, generateSummary]);


  const handleFile = (file: File) => {
    if (processedFiles.includes(file.name)) {
        setError(`تم تحميل هذا الملف "${file.name}" مسبقًا.`);
        return;
    }

    setIsLoading(true);
    setError(null);
    setNotification(null);

    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const fileData = e.target?.result;
        const workbook = XLSX.read(fileData, { type: 'binary', cellDates: true });
        
        let allRawData: DataRow[] = [];
        workbook.SheetNames.forEach(sheetName => {
          const worksheet = workbook.Sheets[sheetName];
          const sheetData: DataRow[] = XLSX.utils.sheet_to_json(worksheet, { raw: true, defval: null });
          allRawData = allRawData.concat(sheetData);
        });

        if (allRawData.length > 0) {
          const { cleanData, cleanHeaders } = processAndCleanData(allRawData);
          
          if(cleanData.length === 0){
             throw new Error("لم يتم العثور على بيانات صالحة في الملف بعد المعالجة.");
          }
          
            const existingRecordKeys = new Set(
              data.map(row => `${row['الاسم']}|${row['التاريخ']}|${row['نوع الاجازة']}|${String(row['القيمة'] ?? '')}`)
            );
            
            const newUniqueRecords = cleanData.filter(row => {
              const key = `${row['الاسم']}|${row['التاريخ']}|${row['نوع الاجازة']}|${String(row['القيمة'] ?? '')}`;
              if (existingRecordKeys.has(key)) {
                return false;
              }
              existingRecordKeys.add(key); // Handle duplicates within the same file
              return true;
            });
    
            const duplicateCount = cleanData.length - newUniqueRecords.length;
            
            if (newUniqueRecords.length > 0) {
                setData(prevData => [...prevData, ...newUniqueRecords]);
                let message = `تمت معالجة الملف "${file.name}". أُضيفت ${toArabicNumerals(newUniqueRecords.length)} سجلات جديدة.`;
                if (duplicateCount > 0) {
                    message += ` وتم تجاهل ${toArabicNumerals(duplicateCount)} سجلات مكررة.`;
                }
                setNotification(message);
                setTimeout(() => setNotification(null), 7000);
            } else {
                 setError(`الملف "${file.name}" لم يضف أي سجلات جديدة. تم تجاهل ${toArabicNumerals(duplicateCount)} سجلات لكونها مكررة.`);
            }

          setHeaders(cleanHeaders);
          setProcessedFiles(prev => [...prev, file.name]);
          setView('table');
        } else {
            setError("الملف فارغ أو لا يحتوي على بيانات.");
        }
      } catch (err) {
        console.error("Error processing Excel file:", err);
        const errorMessage = err instanceof Error ? err.message : "حدث خطأ أثناء معالجة الملف. يرجى التأكد من أنه ملف Excel صالح.";
        setError(errorMessage);
      } finally {
        setIsLoading(false);
      }
    };
    reader.onerror = () => {
        setError("فشل في قراءة الملف.");
        setIsLoading(false);
    }
    reader.readAsBinaryString(file);
  };

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) handleFile(file);
    e.target.value = '';
  };
  
  const handleDragOver = (e: React.DragEvent<HTMLDivElement>) => e.preventDefault();
  const handleDrop = (e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    const file = e.dataTransfer.files?.[0];
    if (file) handleFile(file);
  };

  const requestSort = (key: string) => {
    let direction: 'ascending' | 'descending' = 'ascending';
    if (sortConfig?.key === key && sortConfig.direction === 'ascending') {
      direction = 'descending';
    }
    setSortConfig({ key, direction });
  };

  const processedData = useMemo(() => {
    let filteredData = [...data];
    if (searchTerm) {
      filteredData = filteredData.filter(row =>
        Object.values(row).some(value =>
          String(value).toLowerCase().includes(searchTerm.toLowerCase())
        )
      );
    }
    if (sortConfig !== null) {
      filteredData.sort((a, b) => {
        const valA = a[sortConfig.key];
        const valB = b[sortConfig.key];
        if (valA == null || valA === '') return 1;
        if (valB == null || valB === '') return -1;
        if (valA < valB) return sortConfig.direction === 'ascending' ? -1 : 1;
        if (valA > valB) return sortConfig.direction === 'ascending' ? 1 : -1;
        return 0;
      });
    }
    return filteredData;
  }, [data, searchTerm, sortConfig]);

  const handleDownloadCSV = useCallback(() => {
    if (processedData.length === 0) return;
    const worksheet = XLSX.utils.json_to_sheet(processedData);
    const csvOutput: string = XLSX.utils.sheet_to_csv(worksheet);
    const blob = new Blob(['\uFEFF' + csvOutput], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    const url = URL.createObjectURL(blob);
    link.setAttribute('href', url);
    const fileName = `data_export_${new Date().toISOString().split('T')[0]}.csv`;
    link.setAttribute('download', fileName);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  }, [processedData]);

  const handleClearAllData = () => {
    if (window.confirm("هل أنت متأكد من حذف جميع البيانات؟ سيتم مسح بيانات الموظفين وبيانات الإجازات المحملة بشكل دائم.")) {
        localStorage.removeItem(EMPLOYEES_STORAGE_KEY);
        localStorage.removeItem(LEAVE_DATA_STORAGE_KEY);
        localStorage.removeItem(PROCESSED_FILES_STORAGE_KEY);
        localStorage.removeItem(LAST_BALANCE_UPDATE_KEY);
        window.location.reload();
    }
  };

  const handlePrint = () => {
    window.print();
  };

  const handleExportPDF = async () => {
    if (isLoading) return;
    setIsLoading(true);
    const { jsPDF } = jspdf;
    const outputFileName = `export_${new Date().toISOString().split('T')[0]}.pdf`;

    if (view === 'table') {
       const doc = new jsPDF();
       
       doc.addFileToVFS("Amiri-Regular.ttf", AmiriFont);
       doc.addFont("Amiri-Regular.ttf", "Amiri", "normal");
       doc.setFont("Amiri");

       const head = [headers];
       const body = processedData.map(row => headers.map(header => row[header] !== null && row[header] !== undefined ? String(row[header]) : ''));
       
       const title = 'تقرير بيانات الإكسل';
       const pageWidth = doc.internal.pageSize.getWidth();

       (doc as any).autoTable({
           head: head,
           body: body,
           styles: { font: "Amiri", halign: 'right' },
           headStyles: { halign: 'right', fillColor: [22, 160, 133] },
           didDrawPage: (data: any) => {
               if (data.pageNumber === 1) {
                    doc.setFont("Amiri", "normal");
                    doc.text(title, pageWidth - data.settings.margin.right, 15, { align: 'right' });
               }
           }
       });

       doc.save(outputFileName);

    } else if (view === 'summary' || view === 'rankedSummary') {
        const summaryElement = summaryContentRef.current;
        if (summaryElement) {
            try {
                const canvas = await html2canvas(summaryElement, {
                    scale: 2,
                    backgroundColor: '#ffffff',
                    width: summaryElement.scrollWidth,
                    height: summaryElement.scrollHeight,
                });
                const imgData = canvas.toDataURL('image/png');
                
                const imgWidth = canvas.width;
                const imgHeight = canvas.height;
                
                const pdf = new jsPDF({
                    orientation: imgWidth > imgHeight ? 'l' : 'p',
                    unit: 'px',
                    format: [imgWidth, imgHeight]
                });

                pdf.addImage(imgData, 'PNG', 0, 0, imgWidth, imgHeight);
                pdf.save(outputFileName);
            } catch (err) {
                console.error("Error generating PDF:", err);
                const errorMessage = err instanceof Error ? err.message : "An unknown error occurred during PDF export.";
                setError(`فشل تصدير PDF: ${errorMessage}`);
            }
        }
    }
    setIsLoading(false);
  };

  // --- Auth Handlers ---
  const handleLogin = (e: React.FormEvent) => {
    e.preventDefault();
    setLoginError(null);

    // Admin login
    if (usernameInput === ADMIN_USERNAME && passwordInput === ADMIN_PASSWORD) {
        setCurrentUser('admin');
        setUsernameInput('');
        setPasswordInput('');
        return;
    }

    // Employee login
    const foundEmployee = employees.find(
        emp => emp.username === usernameInput && emp.password === passwordInput
    );

    if (foundEmployee) {
        setCurrentUser(foundEmployee);
        setUsernameInput('');
        setPasswordInput('');
    } else {
        setLoginError('اسم المستخدم أو كلمة المرور غير صحيحة.');
    }
  };

  const handleLogout = () => {
    setCurrentUser(null);
  };
  
    // --- Employee Management Handlers ---
  const handleAddEmployee = async (e: React.FormEvent) => {
    e.preventDefault();
    const name = newEmployeeName.trim();
    const username = newEmployeeUsername.trim();
    const password = newEmployeePassword.trim();
    const balance = parseInt(newEmployeeBalance, 10);
    const workdayHours = parseInt(newEmployeeWorkdayHours, 10);

    if (!name || !username || !password || isNaN(balance) || isNaN(workdayHours)) {
      alert("يرجى إدخال جميع الحقول بشكل صحيح.");
      return;
    }

    if (employees.some(emp => emp.username.toLowerCase() === username.toLowerCase())) {
        alert("اسم المستخدم هذا موجود بالفعل. يرجى اختيار اسم آخر.");
        return;
    }
    
    let photoData: string | undefined = undefined;
    if (newEmployeePhotoFile) {
        photoData = await processImageFile(newEmployeePhotoFile);
    }

    const newEmployee: EmployeeRecord = {
      id: Date.now().toString(),
      name,
      balance,
      username,
      password,
      photo: photoData,
      workdayHours,
    };

    setEmployees(prev => [...prev, newEmployee].sort((a,b) => a.name.localeCompare(b.name, 'ar')));
    setNewEmployeeName('');
    setNewEmployeeBalance('');
    setNewEmployeeWorkdayHours('7');
    setNewEmployeeUsername(generateRandomUsername());
    setNewEmployeePassword(generateRandomPassword());
    setNewEmployeePhoto(null);
    setNewEmployeePhotoFile(null);
  };

  const handleNewEmployeePhotoChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
        setNewEmployeePhotoFile(file);
        const reader = new FileReader();
        reader.onloadend = () => {
            setNewEmployeePhoto(reader.result as string);
        };
        reader.readAsDataURL(file);
    }
  };

  const handleDeleteEmployee = (id: string) => {
    if (window.confirm("هل أنت متأكد من حذف هذا الموظف؟")) {
      setEmployees(prev => prev.filter(emp => emp.id !== id));
    }
  };

  const handleStartEdit = (employee: EmployeeRecord) => {
    setEditingEmployee({ ...employee });
    setEditingEmployeePhotoFile(null);
  };

  const handleCancelEdit = () => {
    setEditingEmployee(null);
    setEditingEmployeePhotoFile(null);
  };
  
 const handleSaveEdit = async () => {
    if (!editingEmployee || isImageProcessing) return;

    const name = editingEmployee.name.trim();
    const username = editingEmployee.username.trim();
    const password = editingEmployee.password.trim();
    const balance = Number(editingEmployee.balance);
    const workdayHours = Number(editingEmployee.workdayHours);

    if (!name || !username || !password || isNaN(balance) || isNaN(workdayHours)) {
        alert("يرجى إدخال جميع الحقول بشكل صحيح.");
        return;
    }
    
    if (employees.some(emp => emp.id !== editingEmployee.id && emp.username.toLowerCase() === username.toLowerCase())) {
        alert("اسم المستخدم هذا موجود بالفعل. يرجى اختيار اسم آخر.");
        return;
    }

    setIsImageProcessing(true);
    setError(null);
    try {
        let finalEmployeeData = { ...editingEmployee };

        if (editingEmployeePhotoFile) {
            const base64 = await processImageFile(editingEmployeePhotoFile);
            finalEmployeeData.photo = base64;
        }

        setEmployees(prev =>
            prev.map(emp => (emp.id === finalEmployeeData.id ? finalEmployeeData : emp))
            .sort((a,b) => a.name.localeCompare(b.name, 'ar'))
        );
        
        setEditingEmployee(null);
        setEditingEmployeePhotoFile(null);
    } catch (err) {
        console.error("Failed to save employee data:", err);
        setError("فشل حفظ التغييرات. قد تكون هناك مشكلة في معالجة الصورة.");
    } finally {
        setIsImageProcessing(false);
    }
};

const handleEditPhotoChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file || !editingEmployee) return;

    setEditingEmployeePhotoFile(file);

    // Show a temporary preview
    const reader = new FileReader();
    reader.onloadend = () => {
        setEditingEmployee(current => {
            if (!current) return null;
            return { ...current, photo: reader.result as string };
        });
    };
    reader.readAsDataURL(file);
};

  
  const filteredEmployees = useMemo(() => {
    if (!employeeSearchTerm) return employees;
    return employees.filter(emp => 
        emp.name.toLowerCase().includes(employeeSearchTerm.toLowerCase()) ||
        emp.username.toLowerCase().includes(employeeSearchTerm.toLowerCase())
    );
  }, [employees, employeeSearchTerm]);

  const handleExportUserList = useCallback(() => {
    if (employees.length === 0) return;
    const headers = ['الاسم', 'اسم المستخدم', 'كلمة المرور'];
    const csvRows = [
        headers.join(','),
        ...employees.map(emp => [
            `"${emp.name.replace(/"/g, '""')}"`,
            `"${emp.username}"`,
            `"${emp.password}"`,
        ].join(','))
    ];
    const csvContent = csvRows.join('\n');
    const blob = new Blob(['\uFEFF' + csvContent], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    const url = URL.createObjectURL(blob);
    link.setAttribute('href', url);
    const fileName = `user_credentials_${new Date().toISOString().split('T')[0]}.csv`;
    link.setAttribute('download', fileName);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  }, [employees]);
  
  const handleCorrectBalances = () => {
    if (window.confirm("هل أنت متأكد من أنك تريد خصم 5 أيام من رصيد كل موظف؟ لا يمكن التراجع عن هذا الإجراء.")) {
        setEmployees(prevEmployees => 
            prevEmployees.map(emp => ({ ...emp, balance: emp.balance - 5 }))
        );
        alert("تم خصم 5 أيام من رصيد جميع الموظفين.");
    }
  };

  // --- Data for Summary View ---

  const availableYears = useMemo(() => {
    if (data.length === 0) return [];
    const years = new Set<string>();
    data.forEach(row => {
        const date = parseDate(String(row['التاريخ']));
        if (!isNaN(date.getTime())) {
            years.add(String(date.getUTCFullYear()));
        }
    });
    return Array.from(years).sort((a, b) => parseInt(b) - parseInt(a));
  }, [data, parseDate]);

  useEffect(() => {
    if (selectedYear === 'all') {
        setSelectedMonth('all');
    }
  }, [selectedYear]);

  const displayedSummary = useMemo(() => {
      const dateFilteredData = data.filter(row => {
          if (selectedYear === 'all') return true;
          const date = parseDate(String(row['التاريخ']));
          if (isNaN(date.getTime())) return false;

          const year = date.getUTCFullYear();
          const month = date.getUTCMonth() + 1;

          const yearMatches = String(year) === selectedYear;
          if (!yearMatches) return false;

          const monthMatches = selectedMonth === 'all' || String(month) === selectedMonth;
          return monthMatches;
      });
      
      let finalSummary = generateSummary(dateFilteredData, employees);
      
      // If the current user is an employee, filter for their data only
      if (currentUser && currentUser !== 'admin') {
          return finalSummary.filter(emp => emp.name === currentUser.name);
      }
      
      // For admin, apply search term
      if (summarySearchTerm) {
          finalSummary = finalSummary.filter(emp => emp.name.toLowerCase().includes(summarySearchTerm.toLowerCase()));
      }

      return finalSummary;
  }, [data, summarySearchTerm, selectedYear, selectedMonth, generateSummary, employees, parseDate, currentUser]);

  const displayedUniqueLeaveTypes = useMemo(() => {
    const allTypes = new Set<string>();
    displayedSummary.forEach(employee => {
        employee.leaves.forEach(leave => {
            allTypes.add(leave.type);
        });
    });
    return Array.from(allTypes).sort((a, b) => {
        const order = ['اجازة اعتيادية', 'اجازة مرضية', 'ملخص الزمنيات'];
        const indexA = order.indexOf(a);
        const indexB = order.indexOf(b);
        if(indexA > -1 && indexB > -1) return indexA - indexB;
        if(indexA > -1) return -1;
        if(indexB > -1) return 1;
        return a.localeCompare(b, 'ar');
    });
  }, [displayedSummary]);
  
  const getRegularAndTimeBasedTotalDays = useCallback((employee: EmployeeSummary): number => {
    let total = 0;
    employee.leaves.forEach(l => {
        if (l.type === 'اجازة اعتيادية' || l.type === 'ملخص الزمنيات') {
            total += l.dayCount;
        }
    });
    return total;
  }, []);
  
  const getSummaryTotalLeaveDays = useCallback((employee: EmployeeSummary): number => {
    return employee.leaves.reduce((total, leave) => total + leave.dayCount, 0);
  }, []);

  const sortedSummary = useMemo(() => {
    return [...summary].sort((a, b) => {
        const totalA = getRegularAndTimeBasedTotalDays(a);
        const totalB = getRegularAndTimeBasedTotalDays(b);
        return totalA - totalB;
    });
  }, [summary, getRegularAndTimeBasedTotalDays]);


  const rankedFilteredSummary = useMemo(() => {
    return sortedSummary.filter(emp => getRegularAndTimeBasedTotalDays(emp) >= 1);
  }, [sortedSummary, getRegularAndTimeBasedTotalDays]);
  
  const rankedSummaryLeaveTypes = useMemo(() => 
    displayedUniqueLeaveTypes.filter(type => type !== 'اجازة مرضية' && type !== 'اجازة طويلة'), 
    [displayedUniqueLeaveTypes]
  );

  const rankedSummaryTableHeaders = useMemo(() => 
    ['ت', 'الاسم', 'عدد ايام الاجازات', ...rankedSummaryLeaveTypes.map(type => type === 'ملخص الزمنيات' ? 'عن ساعات زمنية' : type)], 
    [rankedSummaryLeaveTypes]
  );
  
  const alphabeticallySortedSummary = useMemo(() => {
    return [...rankedFilteredSummary].sort((a, b) => a.name.localeCompare(b.name, 'ar'));
  }, [rankedFilteredSummary]);

  const groupedRankedSummary = useMemo(() => {
    const groups: Record<number, EmployeeSummary[]> = {};
    
    rankedFilteredSummary.forEach(employee => {
        const totalDays = getRegularAndTimeBasedTotalDays(employee);
        if (totalDays > 0) {
            if (!groups[totalDays]) {
                groups[totalDays] = [];
            }
            groups[totalDays].push(employee);
        }
    });

    const getTimeSummaryDays = (employee: EmployeeSummary): number => {
        const timeSummary = employee.leaves.find(l => l.type === 'ملخص الزمنيات');
        return timeSummary ? timeSummary.dayCount : 0;
    };

    for (const dayCount in groups) {
        groups[dayCount].sort((a, b) => {
            const timeDaysA = getTimeSummaryDays(a);
            const timeDaysB = getTimeSummaryDays(b);

            if (timeDaysA === 0 && timeDaysB > 0) return -1;
            if (timeDaysB === 0 && timeDaysA > 0) return 1;
            
            return a.name.localeCompare(b.name, 'ar');
        });
    }

    return groups;
  }, [rankedFilteredSummary, getRegularAndTimeBasedTotalDays]);

  const handleCopyToClipboard = useCallback(() => {
    if (displayedSummary.length === 0) return;

    const headers = ['الاسم', ...displayedUniqueLeaveTypes, 'مجموع الاجازات الاعتيادية', 'المجموع الكلي'];
    let textToCopy = headers.join('\t') + '\n';
    
    displayedSummary.forEach(employee => {
        const leaveMap = new Map<string, LeaveSummary[]>();
        employee.leaves.forEach(l => {
            if (!leaveMap.has(l.type)) leaveMap.set(l.type, []);
            leaveMap.get(l.type)!.push(l);
        });
        
        const rowData = [employee.name];

        displayedUniqueLeaveTypes.forEach(type => {
            const leaves = leaveMap.get(type);
            let content = '';
            if (leaves && leaves.length > 0) {
                const totalDays = leaves.reduce((sum, l) => sum + l.dayCount, 0);
                const totalHours = leaves.reduce((sum, l) => sum + l.hourCount, 0);
                
                const parts: string[] = [];
                if (totalDays > 0) parts.push(`${toArabicNumerals(totalDays)} يوم`);
                if (totalHours > 0) parts.push(`${toArabicNumerals(totalHours)} ساعة`);
                if (parts.length > 0) content = parts.join(' و ');
            }
            rowData.push(content);
        });
        
        const regularAndTimeBasedTotal = getRegularAndTimeBasedTotalDays(employee);
        rowData.push(regularAndTimeBasedTotal > 0 ? formatDaysArabic(regularAndTimeBasedTotal) : '-');

        let totalDays = 0;
        let totalHours = 0;
        employee.leaves.forEach(l => {
            totalDays += l.dayCount;
            totalHours += l.hourCount;
        });
        
        const employeeRecord = employees.find(e => e.name === employee.name);
        const workdayHours = employeeRecord?.workdayHours || 7;

        if (totalHours >= workdayHours) {
            totalDays += Math.floor(totalHours / workdayHours);
            totalHours %= workdayHours;
        }
        const totalParts: string[] = [];
        if (totalDays > 0) totalParts.push(`${toArabicNumerals(totalDays)} يوم`);
        const totalHoursFormatted = parseFloat(totalHours.toFixed(2));
        if (totalHoursFormatted > 0) totalParts.push(`${toArabicNumerals(totalHoursFormatted)} ساعة`);
        const totalContent = totalParts.join(' و ') || '٠';
        rowData.push(totalContent);

        textToCopy += rowData.join('\t') + '\n';
    });


    navigator.clipboard.writeText(textToCopy.trim()).then(() => {
        setShowCopyNotification(true);
        setTimeout(() => setShowCopyNotification(false), 2000); 
    });
  }, [displayedSummary, displayedUniqueLeaveTypes, getRegularAndTimeBasedTotalDays, employees]);

    const handleExportSummaryToWord = useCallback(() => {
        if (displayedSummary.length === 0) return;

        const headers = ['الاسم', ...displayedUniqueLeaveTypes, 'مجموع الاجازات الاعتيادية', 'المجموع الكلي'];
        let tableHTML = `<table border="1" style="border-collapse: collapse; width: 100%; text-align: center; font-family: 'Amiri', Arial, sans-serif;">
            <thead style="background-color: #f2f2f2;">
                <tr>${headers.map(h => `<th style="padding: 8px;">${h}</th>`).join('')}</tr>
            </thead>
            <tbody>`;

        displayedSummary.forEach(employee => {
            const leaveMap = new Map<string, LeaveSummary[]>();
            employee.leaves.forEach(l => {
                if (!leaveMap.has(l.type)) leaveMap.set(l.type, []);
                leaveMap.get(l.type)!.push(l);
            });
            
            tableHTML += '<tr>';
            tableHTML += `<td style="padding: 8px; text-align: right; padding-right: 10px;">${employee.name}</td>`;

            displayedUniqueLeaveTypes.forEach(type => {
                const leaves = leaveMap.get(type);
                let cellContent = '';
                if (leaves && leaves.length > 0) {
                    const totalDays = leaves.reduce((sum, l) => sum + l.dayCount, 0);
                    const totalHours = leaves.reduce((sum, l) => sum + l.hourCount, 0);
                    const parts: string[] = [];
                    if (totalDays > 0) parts.push(`${toArabicNumerals(totalDays)} يوم`);
                    if (totalHours > 0) parts.push(`${toArabicNumerals(totalHours)} ساعة`);
                    if (parts.length > 0) cellContent = parts.join(' و ');
                }
                tableHTML += `<td style="padding: 8px;">${cellContent}</td>`;
            });

            const regularAndTimeBasedTotal = getRegularAndTimeBasedTotalDays(employee);
            tableHTML += `<td style="padding: 8px; font-weight: bold; color: #4338ca;">${regularAndTimeBasedTotal > 0 ? formatDaysArabic(regularAndTimeBasedTotal) : '-'}</td>`;
            
            let totalDays = 0;
            let totalHours = 0;
            employee.leaves.forEach(l => {
                totalDays += l.dayCount;
                totalHours += l.hourCount;
            });
            
            const employeeRecord = employees.find(e => e.name === employee.name);
            const workdayHours = employeeRecord?.workdayHours || 7;

            if (totalHours >= workdayHours) {
                totalDays += Math.floor(totalHours / workdayHours);
                totalHours %= workdayHours;
            }
            const totalParts: string[] = [];
            if (totalDays > 0) totalParts.push(`${toArabicNumerals(totalDays)} يوم`);
            const totalHoursFormatted = parseFloat(totalHours.toFixed(2));
            if (totalHoursFormatted > 0) totalParts.push(`${toArabicNumerals(totalHoursFormatted)} ساعة`);
            const totalContent = totalParts.join(' و ') || '٠';

            tableHTML += `<td style="padding: 8px; font-weight: bold;">${totalContent}</td>`;
            tableHTML += '</tr>';
        });


        tableHTML += `</tbody></table>`;

        const htmlContent = `
            <html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>
            <head><meta charset='utf-8'><title>Export HTML To Doc</title></head>
            <body dir="rtl">
                <h1 style="text-align: center; font-family: 'Amiri', Arial, sans-serif;">ملخص إجازات الموظفين</h1>
                ${tableHTML}
            </body></html>
        `;

        const blob = new Blob(['\uFEFF' + htmlContent], { type: 'application/msword' });
        const url = URL.createObjectURL(blob);
        const link = document.createElement("a");
        link.href = url;
        link.download = `ملخص_إجازات_الموظفين.doc`;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        URL.revokeObjectURL(url);
    }, [displayedSummary, displayedUniqueLeaveTypes, getRegularAndTimeBasedTotalDays, employees]);
  
  const handleExportWord = useCallback(() => {
    let allTablesHTML = '';
    const getLeaveMap = (employee: EmployeeSummary) => {
        const map = new Map<string, LeaveSummary[]>();
        employee.leaves.forEach(l => {
            if (!map.has(l.type)) map.set(l.type, []);
            map.get(l.type)!.push(l);
        });
        return map;
    };


    if (rankedSortOrder === 'alphabetical') {
        allTablesHTML += `<h2 style="text-align: center; font-family: 'Times New Roman'; font-size: 16pt; margin-top: 20px;">اسماء الموظفين مرتبة أبجديًا</h2>`;
        
        let tableHTML = `<table border="1" style="border-collapse: collapse; width: 100%; text-align: center; font-family: 'Times New Roman'; font-size: 14pt;">
            <thead style="background-color: #f2f2f2;">
                <tr>${rankedSummaryTableHeaders.map(h => `<th style="padding: 8px; font-size: 16pt; font-family: 'Times New Roman';">${h}</th>`).join('')}</tr>
            </thead>
            <tbody>`;

        alphabeticallySortedSummary.forEach((employee, index) => {
            const totalDays = getRegularAndTimeBasedTotalDays(employee);
            const leaveMap = getLeaveMap(employee);

            tableHTML += '<tr>';
            tableHTML += `<td style="padding: 8px; font-size: 14pt; font-family: 'Times New Roman';">${toArabicNumerals(index + 1)}</td>`;
            tableHTML += `<td style="padding: 8px; font-size: 14pt; font-family: 'Times New Roman'; text-align: right;">${employee.name}</td>`;
            tableHTML += `<td style="padding: 8px; font-size: 14pt; font-family: 'Times New Roman';">${formatDaysArabic(totalDays)}</td>`;
            
            rankedSummaryLeaveTypes.forEach(type => {
                const leaves = leaveMap.get(type) || [];
                let content = ' ';
                if (leaves.length > 0) {
                    if (type === 'ملخص الزمنيات') {
                        const dayCount = leaves.reduce((sum, l) => sum + l.dayCount, 0);
                        if (dayCount > 0) content = `(${toArabicNumerals(dayCount)}) يوم`;
                    } else if (type === 'اجازة اعتيادية') {
                        const dateDetails = leaves.map(l => l.dateDetails).join(' | ');
                        if (dateDetails) {
                           const datePart = dateDetails.split('|').map(part => {
                                const segments = part.trim().split('/');
                                if (segments.length === 3) {
                                    const days = segments[0].split('،').map(d => parseInt(d.replace(/[٠-٩]/g, d => '٠١٢٣٤٥٦٧٨٩'.indexOf(d).toString()))).sort((a,b) => a-b).map(d=> toArabicNumerals(d)).join('،');
                                    return `${days}/${segments[1]}/${segments[2]}`;
                                }
                                return part;
                            }).join(' | ');
                            content = `<span dir="ltr" style="unicode-bidi: embed;">${datePart}</span>`;
                        } else {
                            content = ' ';
                        }
                    }
                }
                tableHTML += `<td style="padding: 8px; font-size: 14pt; font-family: 'Times New Roman';">${content}</td>`;
            });
            
            tableHTML += '</tr>';
        });
        tableHTML += '</tbody></table>';
        allTablesHTML += tableHTML;

    } else {
        const sortedDayCounts = Object.keys(groupedRankedSummary).sort((a, b) => parseInt(a) - parseInt(b));
        sortedDayCounts.forEach(dayCount => {
            const dayCountNum = parseInt(dayCount, 10);
            allTablesHTML += `<h2 style="text-align: center; font-family: 'Times New Roman'; font-size: 16pt; margin-top: 20px;">اسماء الموظفين الذين تم منحهم إجازة اعتيادية لمدة ${formatDaysArabic(dayCountNum)}</h2>`;
            
            let tableHTML = `<table border="1" style="border-collapse: collapse; width: 100%; text-align: center; font-family: 'Times New Roman'; font-size: 14pt;">
                <thead style="background-color: #f2f2f2;">
                    <tr>${rankedSummaryTableHeaders.map(h => `<th style="padding: 8px; font-size: 16pt; font-family: 'Times New Roman';">${h}</th>`).join('')}</tr>
                </thead>
                <tbody>`;

            const employeesInGroup = groupedRankedSummary[dayCountNum];
            employeesInGroup.forEach((employee, index) => {
                const leaveMap = getLeaveMap(employee);
                
                tableHTML += '<tr>';
                tableHTML += `<td style="padding: 8px; font-size: 14pt; font-family: 'Times New Roman';">${toArabicNumerals(index + 1)}</td>`;
                tableHTML += `<td style="padding: 8px; font-size: 14pt; font-family: 'Times New Roman'; text-align: right;">${employee.name}</td>`;
                tableHTML += `<td style="padding: 8px; font-size: 14pt; font-family: 'Times New Roman';">${formatDaysArabic(dayCountNum)}</td>`;
                
                rankedSummaryLeaveTypes.forEach(type => {
                    const leaves = leaveMap.get(type) || [];
                    let content = ' ';
                    if (leaves.length > 0) {
                       if (type === 'ملخص الزمنيات') {
                            const dayCount = leaves.reduce((sum, l) => sum + l.dayCount, 0);
                            if (dayCount > 0) content = `(${toArabicNumerals(dayCount)}) يوم`;
                        } else if (type === 'اجازة اعتيادية') {
                           const dateDetails = leaves.map(l => l.dateDetails).join(' | ');
                           if (dateDetails) {
                                const datePart = dateDetails.split('|').map(part => {
                                    const segments = part.trim().split('/');
                                    if (segments.length === 3) {
                                        const days = segments[0].split('،').map(d => parseInt(d.replace(/[٠-٩]/g, d => '٠١٢٣٤٥٦٧٨٩'.indexOf(d).toString()))).sort((a,b) => a-b).map(d=> toArabicNumerals(d)).join('،');
                                        return `${days}/${segments[1]}/${segments[2]}`;
                                    }
                                    return part;
                                }).join(' | ');
                                content = `<span dir="ltr" style="unicode-bidi: embed;">${datePart}</span>`;
                            } else {
                                content = ' ';
                            }
                        }
                    }
                    tableHTML += `<td style="padding: 8px; font-size: 14pt; font-family: 'Times New Roman';">${content}</td>`;
                });

                tableHTML += '</tr>';
            });
            tableHTML += '</tbody></table>';
            allTablesHTML += tableHTML;
        });
    }

    const shortSickLeaves = summary
        .flatMap(emp => 
            emp.leaves
                .filter(l => l.type.includes('مرضية') && l.dayCount <= 5)
                .map(l => ({ name: emp.name, ...l }))
        )
        .sort((a,b) => a.dayCount - b.dayCount || a.name.localeCompare(b.name, 'ar'));

    if (shortSickLeaves.length > 0) {
        let sickLeaveTableHTML = `<h2 style="text-align: center; font-family: 'Times New Roman'; font-size: 16pt; margin-top: 20px;">اسماء الموظفين الذين تم منحهم إجازة مرضية</h2>`;
        sickLeaveTableHTML += `<table border="1" style="border-collapse: collapse; width: 100%; text-align: center; font-family: 'Times New Roman'; font-size: 14pt;">
            <thead style="background-color: #f2f2f2;">
                <tr>
                    <th style="padding: 8px; font-size: 16pt; font-family: 'Times New Roman';">ت</th>
                    <th style="padding: 8px; font-size: 16pt; font-family: 'Times New Roman';">الاسم</th>
                    <th style="padding: 8px; font-size: 16pt; font-family: 'Times New Roman';">عدد ايام الاجازة</th>
                    <th style="padding: 8px; font-size: 16pt; font-family: 'Times New Roman';">تاريخ الاجازة من - الى</th>
                </tr>
            </thead>
            <tbody>`;

        shortSickLeaves.forEach((leave, index) => {
            sickLeaveTableHTML += '<tr>';
            sickLeaveTableHTML += `<td style="padding: 8px; font-size: 14pt; font-family: 'Times New Roman';">${toArabicNumerals(index + 1)}</td>`;
            sickLeaveTableHTML += `<td style="padding: 8px; font-size: 14pt; font-family: 'Times New Roman'; text-align: right;">${leave.name}</td>`;
            sickLeaveTableHTML += `<td style="padding: 8px; font-size: 14pt; font-family: 'Times New Roman';">${formatDaysArabic(leave.dayCount)}</td>`;
            sickLeaveTableHTML += `<td style="padding: 8px; font-size: 14pt; font-family: 'Times New Roman';">${leave.dateDetails}</td>`;
            sickLeaveTableHTML += '</tr>';
        });

        sickLeaveTableHTML += '</tbody></table>';
        allTablesHTML += sickLeaveTableHTML;
    }


    const htmlContent = `
        <html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>
        <head><meta charset='utf-8'><title>Export HTML To Doc</title></head>
        <body dir="rtl">
            ${allTablesHTML}
        </body></html>
    `;

    const blob = new Blob(['\uFEFF' + htmlContent], { type: 'application/msword' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = `ملخص_مرتب.doc`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
}, [summary, groupedRankedSummary, rankedSummaryTableHeaders, rankedSortOrder, alphabeticallySortedSummary, getRegularAndTimeBasedTotalDays, rankedSummaryLeaveTypes]);

    // --- Monthly Report Handlers ---
    const handleGenerateMonthlyReport = useCallback(() => {
        setIsLoading(true);
        const reportStartDate = new Date(Date.UTC(reportYear, reportMonth - 1, 1));
        const reportEndDate = new Date(Date.UTC(reportYear, reportMonth, 0));

        const normalizeName = (name: string): string => {
            if (!name) return '';
            return name.trim().replace(/\s+/g, '').replace(/[أإآ]/g, 'ا').replace(/ة/g, 'ه').replace(/ى/g, 'ي');
        };
        const canonicalNameMap = new Map(employees.map(e => [normalizeName(e.name), e.name]));

        const resolveCanonicalName = (sheetName: string): string => {
            const normalized = normalizeName(sheetName);
            return canonicalNameMap.get(normalized) || sheetName;
        };

        const newReportData: MonthlyReportRow[] = employees.map(employee => {
            const leavesForEmployee = data.filter(d => resolveCanonicalName(d['الاسم']) === employee.name);
            
            let workdayHours = employee.workdayHours || 7; // Start with stored or default
            // Dynamically determine workday hours from actual data, overriding stored value
            const regularLeaves = leavesForEmployee.filter(l => String(l['نوع الاجازة']).includes('اعتيادية'));
            if (regularLeaves.length > 0) {
                const typicalHours = parseFloat(String(regularLeaves[0]['القيمة']));
                if (typicalHours === 6 || typicalHours === 7) {
                    workdayHours = typicalHours;
                }
            }


            // Leaves within the current month
            const monthlyLeaves = leavesForEmployee.filter(row => {
                const date = parseDate(String(row['التاريخ']));
                return date >= reportStartDate && date <= reportEndDate;
            });

            const regularLeavesMonth = monthlyLeaves.filter(l => String(l['نوع الاجازة']).includes('اعتيادية'));
            const regularDates = regularLeavesMonth.map(l => parseDate(String(l['التاريخ'])).getUTCDate()).sort((a,b)=>a-b).map(toArabicNumerals).join('، ');

            const getLeaveDateRange = (leaveTypeKeyword: string): string => {
                const leaves = monthlyLeaves.filter(l => String(l['نوع الاجازة']).includes(leaveTypeKeyword));
                if (leaves.length === 0) return '';
                const dates = leaves.map(l => parseDate(String(l['التاريخ']))).sort((a, b) => a.getTime() - b.getTime());
                const minDate = dates[0];
                const maxDate = dates[dates.length - 1];
                if (minDate.getTime() === maxDate.getTime()) {
                    const d = minDate.getUTCDate();
                    const m = minDate.getUTCMonth() + 1;
                    const y = minDate.getUTCFullYear();
                    return toArabicNumerals(`${d}/${m}/${y}`);
                }

                const startDay = minDate.getUTCDate();
                const endDay = maxDate.getUTCDate();
                const year = minDate.getUTCFullYear();
                return toArabicNumerals(`${endDay}-${startDay}/${year}`);
            };

            const sickLeaveDateRange = getLeaveDateRange('مرضية');
            const longLeaveDateRange = getLeaveDateRange('طويلة');
            
            // --- Cumulative Hourly Calculation for monthly display ---
            const leavesBeforeMonth = leavesForEmployee.filter(row => parseDate(String(row['التاريخ'])) < reportStartDate);
            
            const hoursBeforeMonthSheet = leavesBeforeMonth
                .filter(l => String(l['نوع الاجازة']).startsWith('زمنية'))
                .reduce((sum, l) => sum + (parseFloat(String(l['القيمة'])) || 0), 0);
            const totalHoursBeforeMonth = hoursBeforeMonthSheet + (employee.priorHourlyBalance || 0);

            const daysBeforeMonth = Math.floor(totalHoursBeforeMonth / workdayHours);

            const hourlyLeavesMonth = monthlyLeaves.filter(l => String(l['نوع الاجازة']).startsWith('زمنية'));
            const hoursThisMonth = hourlyLeavesMonth.reduce((sum, l) => sum + (parseFloat(String(l['القيمة'])) || 0), 0);
            
            const totalHoursAtMonthEnd = totalHoursBeforeMonth + hoursThisMonth;
            const daysAtMonthEnd = Math.floor(totalHoursAtMonthEnd / workdayHours);
            const remainingHoursAtMonthEnd = totalHoursAtMonthEnd % workdayHours;

            const daysToShowForMonth = daysAtMonthEnd - daysBeforeMonth;
            const hoursToShowForMonth = remainingHoursAtMonthEnd;

            // --- Calculation for final remaining balance ---
            const leavesUntilMonthEnd = leavesForEmployee.filter(row => {
                const date = parseDate(String(row['التاريخ']));
                return date <= reportEndDate;
            });

            const totalRegularDays = leavesUntilMonthEnd.filter(l => String(l['نوع الاجازة']).includes('اعتيادية')).length;
            
            // Note: total hourly days deducted is simply `daysAtMonthEnd`
            const finalBalance = (employee.balance || 0) - totalRegularDays - daysAtMonthEnd;

            return {
                name: employee.name,
                initialBalance: employee.balance || 0,
                regularLeaves: { count: regularLeavesMonth.length, dates: regularDates },
                hourlyLeaves: { days: daysToShowForMonth, hours: hoursToShowForMonth },
                sickLeave: { dateRange: sickLeaveDateRange },
                longLeave: { dateRange: longLeaveDateRange },
                finalBalance: finalBalance
            };
        });

        setMonthlyReportData(newReportData);
        setIsLoading(false);
    }, [reportYear, reportMonth, data, employees, parseDate]);

    const handleExportMonthlyReportToWord = useCallback(() => {
        if (monthlyReportData.length === 0) return;

        const headers = ['ت', 'اسم الموظف', 'الرصيد الأولي', 'الاعتيادية وتاريخها', 'الزمنية (أيام/ساعات)', 'المرضية', 'الطويلة', 'الرصيد المتبقي'];
        let tableHTML = `<table border="1" style="border-collapse: collapse; width: 100%; text-align: center; font-family: 'Times New Roman'; font-size: 14pt;">
            <thead style="background-color: #f2f2f2;">
                <tr>${headers.map(h => `<th style="padding: 8px; font-size: 16pt; font-family: 'Times New Roman';">${h}</th>`).join('')}</tr>
            </thead>
            <tbody>`;

        monthlyReportData.forEach((row, index) => {
            const regularLeaveText = row.regularLeaves.count > 0 ? `${formatDaysArabic(row.regularLeaves.count)} <span dir="ltr">/${row.regularLeaves.dates}</span>` : '-';
            const hourlyLeaveText = (row.hourlyLeaves.days > 0 || row.hourlyLeaves.hours > 0) ? `${toArabicNumerals(row.hourlyLeaves.days)} يوم / ${toArabicNumerals(parseFloat(row.hourlyLeaves.hours.toFixed(2)))} س` : '-';

            tableHTML += '<tr>';
            tableHTML += `<td style="padding: 8px;">${toArabicNumerals(index + 1)}</td>`;
            tableHTML += `<td style="padding: 8px; text-align: right;">${row.name}</td>`;
            tableHTML += `<td style="padding: 8px;">${toArabicNumerals(row.initialBalance)}</td>`;
            tableHTML += `<td style="padding: 8px;">${regularLeaveText}</td>`;
            tableHTML += `<td style="padding: 8px;">${hourlyLeaveText}</td>`;
            tableHTML += `<td style="padding: 8px;">${row.sickLeave.dateRange || '-'}</td>`;
            tableHTML += `<td style="padding: 8px;">${row.longLeave.dateRange || '-'}</td>`;
            tableHTML += `<td style="padding: 8px; font-weight: bold;">${toArabicNumerals(row.finalBalance)}</td>`;
            tableHTML += '</tr>';
        });

        tableHTML += '</tbody></table>';
        
        const monthName = new Date(reportYear, reportMonth-1, 1).toLocaleString('ar', { month: 'long' });
        const title = `التقرير النهائي لشهر ${monthName} ${toArabicNumerals(reportYear)}`;

        const htmlContent = `
            <html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>
            <head><meta charset='utf-8'><title>${title}</title></head>
            <body dir="rtl">
                <h1 style="text-align: center; font-family: 'Times New Roman'; font-size: 16pt;">${title}</h1>
                ${tableHTML}
            </body></html>
        `;

        const blob = new Blob(['\uFEFF' + htmlContent], { type: 'application/msword' });
        const url = URL.createObjectURL(blob);
        const link = document.createElement("a");
        link.href = url;
        link.download = `${title}.doc`;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        URL.revokeObjectURL(url);
    }, [monthlyReportData, reportMonth, reportYear]);

    // --- Summary View Helpers ---
    const getInitials = (name: string): string => {
        if (!name) return '?';
        return name.split(' ').map(n => n[0]).filter(Boolean).slice(0, 2).join('').toUpperCase();
    };

    const LeaveTypeIcon: React.FC<{ type: string; className?: string }> = ({ type, className = "w-6 h-6" }) => {
        if (type.includes('اعتيادية')) return <BriefcaseIcon className={className} />;
        if (type.includes('مرضية')) return <HeartIcon className={className} />;
        if (type.includes('زمنية')) return <ClockIcon className={className} />;
        return <CalendarDaysIcon className={className} />;
    };

    // Card component for the new summary view
    const EmployeeSummaryCard: React.FC<{ employee: EmployeeSummary }> = ({ employee }) => {
        const regularAndTimeBasedTotal = getRegularAndTimeBasedTotalDays(employee);
        
        const remainingBalance = employee.initialBalance != null 
            ? employee.initialBalance - regularAndTimeBasedTotal 
            : null;

        const regularLeaves = employee.leaves.filter(l => l.type === 'اجازة اعتيادية');
        const regularLeaveDays = regularLeaves.reduce((sum, l) => sum + l.dayCount, 0);

        const timeSummary = employee.leaves.find(l => l.type === 'ملخص الزمنيات');

        const timeSummaryDays = timeSummary?.dayCount || 0;
        const timeSummaryHours = timeSummary?.hourCount || 0;

        const hasFooterData = regularAndTimeBasedTotal > 0 || remainingBalance != null;

        return (
            <div className="bg-white rounded-xl shadow-md hover:shadow-xl transition-shadow duration-300 flex flex-col print:shadow-none print:border print:border-gray-300">
                <div className="p-4 flex items-center gap-4 border-b">
                    {employee.photo ? (
                        <img src={employee.photo} alt={employee.name} className="w-12 h-12 rounded-full object-cover" />
                    ) : (
                        <div className="w-12 h-12 rounded-full bg-indigo-100 text-indigo-600 flex items-center justify-center text-xl font-bold">
                           <UserCircleIcon className="w-8 h-8"/>
                        </div>
                    )}
                    <div>
                        <h3 className="text-lg font-bold text-gray-800">{employee.name}</h3>
                        {employee.initialBalance != null && (
                            <p className="text-sm text-gray-500">الرصيد الأولي: {toArabicNumerals(employee.initialBalance)} يوم</p>
                        )}
                    </div>
                </div>
                <div className="p-4 space-y-3 flex-grow">
                    {employee.leaves.length > 0 ? (
                        employee.leaves.map((leave, index) => (
                            <div key={index} className="flex items-start gap-3">
                                <div className="flex-shrink-0 text-gray-400 pt-1">
                                    <LeaveTypeIcon type={leave.type} />
                                </div>
                                <div>
                                    <p className="font-semibold text-gray-700">{leave.type}</p>
                                    <p className="text-sm text-gray-600">
                                        <span className="font-medium text-indigo-600">{formatLeaveCount(leave.dayCount, leave.hourCount)}</span>
                                        {leave.dateDetails && <span className="text-xs text-gray-400 block">{leave.dateDetails}</span>}
                                    </p>
                                </div>
                            </div>
                        ))
                    ) : (
                        <p className="text-gray-500 text-center py-4">لا توجد إجازات مسجلة.</p>
                    )}
                </div>
                {hasFooterData && (
                     <div className="p-4 bg-gray-50 rounded-b-xl border-t mt-auto">
                        <div className="space-y-1 mb-3 text-sm">
                            {regularLeaveDays > 0 && (
                                <div className="flex justify-between">
                                    <span className="text-gray-600">أيام اعتيادية:</span>
                                    <span className="font-medium text-gray-800">{formatDaysArabic(regularLeaveDays)}</span>
                                </div>
                            )}
                            {timeSummaryDays > 0 && (
                                <div className="flex justify-between">
                                    <span className="text-gray-600">أيام من الزمنيات:</span>
                                    <span className="font-medium text-gray-800">{formatDaysArabic(timeSummaryDays)}</span>
                                </div>
                            )}
                             {timeSummaryHours > 0 && (
                                <div className="flex justify-between">
                                    <span className="text-gray-600">ساعات زمنية متبقية:</span>
                                    <span className="font-medium text-gray-800">{toArabicNumerals(timeSummaryHours)} ساعة</span>
                                </div>
                            )}
                        </div>

                        {(regularLeaveDays > 0 || timeSummaryDays > 0 || timeSummaryHours > 0) && (
                           <hr className="my-2 border-gray-200" />
                        )}

                        <div className="flex justify-between items-center text-sm">
                            <span className="font-semibold text-gray-600">مجموع الاعتيادي والزمني:</span>
                            <span className="font-bold text-lg text-red-600">{formatDaysArabic(regularAndTimeBasedTotal)}</span>
                        </div>
                        {remainingBalance != null && (
                            <div className="flex justify-between items-center mt-2 text-sm">
                                <span className="font-semibold text-gray-600">الرصيد المتبقي:</span>
                                <span className={`font-bold text-lg ${remainingBalance < 0 ? 'text-red-700' : 'text-green-600'}`}>{toArabicNumerals(remainingBalance)} يوم</span>
                            </div>
                        )}
                    </div>
                )}
            </div>
        );
    };

    const renderRankedSummary = () => {
        const leaveMapCache = new WeakMap<EmployeeSummary, Map<string, LeaveSummary[]>>();
        const getLeaveMap = (employee: EmployeeSummary) => {
            if (!leaveMapCache.has(employee)) {
                const map = new Map<string, LeaveSummary[]>();
                employee.leaves.forEach(l => {
                    if (!map.has(l.type)) map.set(l.type, []);
                    map.get(l.type)!.push(l);
                });
                leaveMapCache.set(employee, map);
            }
            return leaveMapCache.get(employee)!;
        };
        
        const shortSickLeaves = summary
            .flatMap(emp => 
                emp.leaves
                    .filter(l => l.type.includes('مرضية') && l.dayCount <= 5)
                    .map(l => ({ name: emp.name, ...l }))
            )
            .sort((a,b) => a.dayCount - b.dayCount || a.name.localeCompare(b.name, 'ar'));

        const sickLeaveTableComponent = shortSickLeaves.length > 0 ? (
            <div className="mt-8">
                <h2 className="text-xl font-bold text-center mb-4">
                    اسماء الموظفين الذين تم منحهم إجازة مرضية
                </h2>
                <table className="w-full text-sm text-center text-gray-600 border border-gray-300">
                    <thead className="text-xs text-gray-700 uppercase bg-gray-100">
                        <tr>
                            <th className="px-2 py-3 border">ت</th>
                            <th className="px-2 py-3 border">الاسم</th>
                            <th className="px-2 py-3 border">عدد ايام الاجازة</th>
                            <th className="px-2 py-3 border">تاريخ الاجازة من - الى</th>
                        </tr>
                    </thead>
                    <tbody>
                        {shortSickLeaves.map((leave, index) => (
                            <tr key={`${leave.name}-${index}`} className="bg-white border-b hover:bg-gray-50">
                                <td className="px-2 py-2 border">{toArabicNumerals(index + 1)}</td>
                                <td className="px-2 py-2 border text-right font-semibold">{leave.name}</td>
                                <td className="px-2 py-2 border">{formatDaysArabic(leave.dayCount)}</td>
                                <td className="px-2 py-2 border">{leave.dateDetails}</td>
                            </tr>
                        ))}
                    </tbody>
                </table>
            </div>
        ) : null;

        let mainContent;

        if (rankedSortOrder === 'alphabetical') {
             if (alphabeticallySortedSummary.length === 0) {
                mainContent = <p className="text-center text-gray-500 mt-8">.لا توجد بيانات لعرضها</p>;
             } else {
                mainContent = (
                    <div>
                         <h2 className="text-xl font-bold text-center mb-4">اسماء الموظفين مرتبة أبجديًا</h2>
                         <table className="w-full text-sm text-center text-gray-600 border border-gray-300">
                            <thead className="text-xs text-gray-700 uppercase bg-gray-100">
                                <tr>
                                    {rankedSummaryTableHeaders.map(h => <th key={h} className="px-2 py-3 border">{h}</th>)}
                                </tr>
                            </thead>
                            <tbody>
                                {alphabeticallySortedSummary.map((employee, index) => {
                                    const totalDays = getRegularAndTimeBasedTotalDays(employee);
                                    const leaveMap = getLeaveMap(employee);
        
                                    return (
                                        <tr key={employee.name} className="bg-white border-b hover:bg-gray-50">
                                            <td className="px-2 py-2 border">{toArabicNumerals(index + 1)}</td>
                                            <td className="px-2 py-2 border text-right font-semibold">{employee.name}</td>
                                            <td className="px-2 py-2 border">{formatDaysArabic(totalDays)}</td>
                                            {rankedSummaryLeaveTypes.map(type => {
                                                const leaves = leaveMap.get(type) || [];
                                                let content = ' ';
                                                if (leaves.length > 0) {
                                                    if (type === 'ملخص الزمنيات') {
                                                        const dayCount = leaves.reduce((sum, l) => sum + l.dayCount, 0);
                                                        if (dayCount > 0) content = `(${toArabicNumerals(dayCount)}) يوم`;
                                                    } else if (type === 'اجازة اعتيادية') {
                                                        content = leaves.map(l => l.dateDetails).join(' | ') || ' ';
                                                    }
                                                }
                                                return <td key={type} className="px-2 py-2 border">{content}</td>;
                                            })}
                                        </tr>
                                    );
                                })}
                            </tbody>
                        </table>
                    </div>
                );
             }
        } else {
            const sortedDayCounts = Object.keys(groupedRankedSummary).sort((a, b) => parseInt(a) - parseInt(b));
            if (sortedDayCounts.length === 0) {
                mainContent = <p className="text-center text-gray-500 mt-8">.لا توجد بيانات لعرضها</p>;
            } else {
                mainContent = (
                    <div className="space-y-8">
                        {sortedDayCounts.map(dayCount => {
                            const dayCountNum = parseInt(dayCount, 10);
                            const employeesInGroup = groupedRankedSummary[dayCountNum];
                            return (
                                <div key={dayCount}>
                                    <h2 className="text-xl font-bold text-center mb-4">
                                        اسماء الموظفين الذين تم منحهم إجازة اعتيادية لمدة {formatDaysArabic(dayCountNum)}
                                    </h2>
                                    <table className="w-full text-sm text-center text-gray-600 border border-gray-300">
                                        <thead className="text-xs text-gray-700 uppercase bg-gray-100">
                                            <tr>
                                                {rankedSummaryTableHeaders.map(h => <th key={h} className="px-2 py-3 border">{h}</th>)}
                                            </tr>
                                        </thead>
                                        <tbody>
                                            {employeesInGroup.map((employee, index) => {
                                                const leaveMap = getLeaveMap(employee);
                                                return (
                                                    <tr key={employee.name} className="bg-white border-b hover:bg-gray-50">
                                                        <td className="px-2 py-2 border">{toArabicNumerals(index + 1)}</td>
                                                        <td className="px-2 py-2 border text-right font-semibold">{employee.name}</td>
                                                        <td className="px-2 py-2 border">{formatDaysArabic(dayCountNum)}</td>
                                                        {rankedSummaryLeaveTypes.map(type => {
                                                            const leaves = leaveMap.get(type) || [];
                                                            let content = ' ';
                                                            if (leaves.length > 0) {
                                                                if (type === 'ملخص الزمنيات') {
                                                                    const dayCount = leaves.reduce((sum, l) => sum + l.dayCount, 0);
                                                                    if (dayCount > 0) content = `(${toArabicNumerals(dayCount)}) يوم`;
                                                                } else if (type === 'اجازة اعتيادية') {
                                                                    content = leaves.map(l => l.dateDetails).join(' | ') || ' ';
                                                                }
                                                            }
                                                            return <td key={type} className="px-2 py-2 border">{content}</td>;
                                                        })}
                                                    </tr>
                                                );
                                            })}
                                        </tbody>
                                    </table>
                                </div>
                            );
                        })}
                    </div>
                );
            }
        }
    
        return (
            <>
                {mainContent}
                {sickLeaveTableComponent}
            </>
        );
    };

    if (!currentUser) {
        return (
          <div className="min-h-screen bg-gray-100 flex items-center justify-center p-4">
            <div className="w-full max-w-md bg-white rounded-lg shadow-xl p-8">
              <h1 className="text-3xl font-bold text-center text-gray-800 mb-2">تسجيل الدخول</h1>
              <p className="text-center text-gray-500 mb-8">الرجاء إدخال بيانات الاعتماد الخاصة بك</p>
              <form onSubmit={handleLogin}>
                <div className="mb-4">
                  <label className="block text-gray-700 text-sm font-bold mb-2" htmlFor="username">
                    اسم المستخدم
                  </label>
                  <input
                    id="username"
                    type="text"
                    value={usernameInput}
                    onChange={(e) => setUsernameInput(e.target.value)}
                    className="shadow-sm appearance-none border rounded w-full py-2 px-3 text-gray-700 leading-tight focus:outline-none focus:ring-2 focus:ring-indigo-500"
                    required
                  />
                </div>
                <div className="mb-6">
                  <label className="block text-gray-700 text-sm font-bold mb-2" htmlFor="password">
                    كلمة المرور
                  </label>
                  <input
                    id="password"
                    type="password"
                    value={passwordInput}
                    onChange={(e) => setPasswordInput(e.target.value)}
                    className="shadow-sm appearance-none border rounded w-full py-2 px-3 text-gray-700 mb-3 leading-tight focus:outline-none focus:ring-2 focus:ring-indigo-500"
                    required
                  />
                </div>
                {loginError && <p className="text-red-500 text-xs italic mb-4">{loginError}</p>}
                <div className="flex items-center justify-between">
                  <button
                    type="submit"
                    className="bg-indigo-600 hover:bg-indigo-700 text-white font-bold py-2 px-4 rounded-lg focus:outline-none focus:shadow-outline transition-colors duration-300 w-full"
                  >
                    تسجيل الدخول
                  </button>
                </div>
              </form>
            </div>
          </div>
        );
    }
    
    // Main App Layout
    return (
        <div className="min-h-screen bg-slate-100 flex">
          {/* Sidebar Navigation */}
          <aside className="w-64 bg-white shadow-lg p-4 flex flex-col no-print">
             <div className="text-center mb-8">
                <h1 className="text-2xl font-bold text-indigo-600">مستخرج الإجازات</h1>
                <p className="text-sm text-gray-500 mt-1">مرحباً, {currentUser === 'admin' ? 'المسؤول' : currentUser.name}</p>
            </div>
            <nav className="flex-grow">
              <ul className="space-y-2">
                 {currentUser === 'admin' && (
                    <li>
                        <button onClick={() => setView('upload')} className={`flex items-center gap-3 w-full p-3 rounded-lg text-right transition-colors ${view === 'upload' ? 'bg-indigo-100 text-indigo-700' : 'text-gray-600 hover:bg-gray-100'}`}>
                            <UploadIcon className="w-6 h-6" />
                            <span>تحميل ملف جديد</span>
                        </button>
                    </li>
                 )}
                <li>
                  <button onClick={() => setView('table')} className={`flex items-center gap-3 w-full p-3 rounded-lg text-right transition-colors ${view === 'table' ? 'bg-indigo-100 text-indigo-700' : 'text-gray-600 hover:bg-gray-100'}`}>
                    <TableCellsIcon className="w-6 h-6" />
                    <span>عرض البيانات الخام</span>
                  </button>
                </li>
                <li>
                  <button onClick={() => setView('summary')} className={`flex items-center gap-3 w-full p-3 rounded-lg text-right transition-colors ${view === 'summary' ? 'bg-indigo-100 text-indigo-700' : 'text-gray-600 hover:bg-gray-100'}`}>
                    <ClipboardDocumentIcon className="w-6 h-6" />
                    <span>ملخص الإجازات</span>
                  </button>
                </li>
                <li>
                    <button onClick={() => setView('rankedSummary')} className={`flex items-center gap-3 w-full p-3 rounded-lg text-right transition-colors ${view === 'rankedSummary' ? 'bg-indigo-100 text-indigo-700' : 'text-gray-600 hover:bg-gray-100'}`}>
                        <TrophyIcon className="w-6 h-6" />
                        <span>الملخص المرتب</span>
                    </button>
                </li>
                 <li>
                    <button onClick={() => setView('monthlyReport')} className={`flex items-center gap-3 w-full p-3 rounded-lg text-right transition-colors ${view === 'monthlyReport' ? 'bg-indigo-100 text-indigo-700' : 'text-gray-600 hover:bg-gray-100'}`}>
                        <DocumentTextIcon className="w-6 h-6" />
                        <span>التقرير الشهري النهائي</span>
                    </button>
                </li>
                {currentUser === 'admin' && (
                  <li>
                    <button onClick={() => setView('employeeManagement')} className={`flex items-center gap-3 w-full p-3 rounded-lg text-right transition-colors ${view === 'employeeManagement' ? 'bg-indigo-100 text-indigo-700' : 'text-gray-600 hover:bg-gray-100'}`}>
                        <UsersIcon className="w-6 h-6" />
                        <span>إدارة الموظفين</span>
                    </button>
                  </li>
                )}
              </ul>
            </nav>
            <div className="mt-auto">
                <button onClick={handleLogout} className="w-full text-left text-gray-600 hover:text-red-600 p-3 rounded-lg hover:bg-red-50 transition-colors">
                    تسجيل الخروج
                </button>
                {currentUser === 'admin' && (
                    <button onClick={handleClearAllData} className="w-full text-left flex items-center gap-2 text-sm text-red-500 hover:text-red-700 p-2 mt-4 rounded-lg hover:bg-red-50 transition-colors">
                        <TrashIcon className="w-4 h-4" />
                        <span>حذف جميع البيانات</span>
                    </button>
                )}
            </div>
          </aside>
    
          {/* Main Content Area */}
          <main className="flex-1 p-6 lg:p-8 overflow-auto printable-area">
            {error && (
              <div className="bg-red-100 border border-red-400 text-red-700 px-4 py-3 rounded-lg relative mb-4 no-print" role="alert">
                <strong className="font-bold">خطأ! </strong>
                <span className="block sm:inline">{error}</span>
                <span className="absolute top-0 bottom-0 left-0 px-4 py-3" onClick={() => setError(null)}>
                  <svg className="fill-current h-6 w-6 text-red-500" role="button" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 20 20"><title>Close</title><path d="M14.348 14.849a1.2 1.2 0 0 1-1.697 0L10 11.819l-2.651 3.029a1.2 1.2 0 1 1-1.697-1.697l2.758-3.15-2.759-3.152a1.2 1.2 0 1 1 1.697-1.697L10 8.183l2.651-3.031a1.2 1.2 0 1 1 1.697 1.697l-2.758 3.152 2.758 3.15a1.2 1.2 0 0 1 0 1.698z"/></svg>
                </span>
              </div>
            )}

            {notification && (
              <div className="bg-green-100 border border-green-400 text-green-700 px-4 py-3 rounded-lg relative mb-4 no-print" role="alert">
                <span className="block sm:inline">{notification}</span>
                <span className="absolute top-0 bottom-0 left-0 px-4 py-3" onClick={() => setNotification(null)}>
                  <svg className="fill-current h-6 w-6 text-green-500" role="button" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 20 20"><title>Close</title><path d="M14.348 14.849a1.2 1.2 0 0 1-1.697 0L10 11.819l-2.651 3.029a1.2 1.2 0 1 1-1.697-1.697l2.758-3.15-2.759-3.152a1.2 1.2 0 1 1 1.697-1.697L10 8.183l2.651-3.031a1.2 1.2 0 1 1 1.697 1.697l-2.758 3.152 2.758 3.15a1.2 1.2 0 0 1 0 1.698z"/></svg>
                </span>
              </div>
            )}
    
            {isLoading && (
                <div className="fixed inset-0 bg-gray-600 bg-opacity-50 flex items-center justify-center z-50">
                    <div className="bg-white p-5 rounded-lg flex items-center space-x-4 space-x-reverse">
                        <div className="animate-spin rounded-full h-12 w-12 border-b-2 border-indigo-700"></div>
                        <p className="text-lg">...جاري التحميل</p>
                    </div>
                </div>
            )}
            
            {showCopyNotification && (
                <div className="fixed top-5 right-5 bg-green-500 text-white py-2 px-4 rounded-lg shadow-lg z-50">
                    تم النسخ إلى الحافظة!
                </div>
            )}
            
            <div className="printable-content">
            {/* Render content based on view */}
            {view === 'upload' && (
                <div 
                  onDragOver={handleDragOver} 
                  onDrop={handleDrop}
                  className="bg-white p-8 rounded-xl shadow-lg border-2 border-dashed border-gray-300 hover:border-indigo-500 transition-colors duration-300 text-center"
                >
                    <div className="mb-4">
                        <UploadIcon className="w-16 h-16 mx-auto text-gray-400" />
                    </div>
                    <h2 className="text-2xl font-bold text-gray-800 mb-2">قم بسحب وإفلات ملف Excel هنا</h2>
                    <p className="text-gray-500 mb-6">أو انقر لتحديد ملف</p>
                    <button
                        onClick={() => fileInputRef.current?.click()}
                        className="bg-indigo-600 hover:bg-indigo-700 text-white font-bold py-2 px-6 rounded-lg transition-colors duration-300"
                        disabled={isLoading}
                    >
                        اختر ملف
                    </button>
                    <input
                        type="file"
                        ref={fileInputRef}
                        onChange={handleFileChange}
                        className="hidden"
                        accept=".xlsx, .xls, .csv"
                    />
                     {processedFiles.length > 0 && (
                        <div className="mt-8 text-right">
                            <h3 className="font-bold text-gray-700 mb-2">الملفات التي تمت معالجتها:</h3>
                            <ul className="list-disc list-inside text-gray-600">
                                {processedFiles.map(name => <li key={name}>{name}</li>)}
                            </ul>
                        </div>
                     )}
                </div>
            )}
    
            {view === 'table' && (
                <div className="bg-white p-6 rounded-xl shadow-lg">
                    <div className="flex flex-col md:flex-row justify-between items-center mb-4 gap-4 no-print">
                        <div className="relative w-full md:w-1/3">
                            <input
                                type="text"
                                placeholder="ابحث في البيانات..."
                                value={searchTerm}
                                onChange={(e) => setSearchTerm(e.target.value)}
                                className="w-full pl-10 pr-4 py-2 border rounded-lg focus:outline-none focus:ring-2 focus:ring-indigo-500"
                            />
                             <SearchIcon className="absolute right-3 top-1/2 transform -translate-y-1/2 w-5 h-5 text-gray-400" />
                        </div>
                        <div className="flex items-center gap-2">
                            <button
                                onClick={handleDownloadCSV}
                                className="bg-green-600 hover:bg-green-700 text-white font-bold py-2 px-4 rounded-lg flex items-center gap-2 transition-colors"
                            >
                                <DownloadIcon className="w-5 h-5"/>
                                <span>CSV تحميل</span>
                            </button>
                             <button onClick={handlePrint} className="bg-sky-500 hover:bg-sky-600 text-white font-bold py-2 px-4 rounded-lg flex items-center gap-2 transition-colors">
                                <PrinterIcon className="w-5 h-5"/>
                                <span>طباعة</span>
                            </button>
                            <button onClick={handleExportPDF} className="bg-red-500 hover:bg-red-600 text-white font-bold py-2 px-4 rounded-lg flex items-center gap-2 transition-colors">
                                <PdfIcon className="w-5 h-5"/>
                                <span>PDF تصدير</span>
                            </button>
                        </div>
                    </div>
    
                    <div className="overflow-x-auto">
                        <table className="w-full text-sm text-right text-gray-500">
                            <thead className="text-xs text-gray-700 uppercase bg-gray-100">
                            <tr>
                                {headers.map((header) => (
                                <th key={header} scope="col" className="px-6 py-3">
                                    <button
                                        onClick={() => requestSort(header)}
                                        className="flex items-center gap-1.5"
                                    >
                                    {header}
                                    <SortIcon direction={sortConfig?.key === header ? sortConfig.direction : undefined} />
                                    </button>
                                </th>
                                ))}
                            </tr>
                            </thead>
                            <tbody>
                            {processedData.map((row, index) => (
                                <tr key={index} className="bg-white border-b hover:bg-gray-50">
                                {headers.map((header) => (
                                    <td key={header} className="px-6 py-4">
                                    {String(row[header] ?? '')}
                                    </td>
                                ))}
                                </tr>
                            ))}
                            </tbody>
                        </table>
                    </div>
                     {processedData.length === 0 && <p className="text-center text-gray-500 mt-8">.لا توجد بيانات لعرضها</p>}
                </div>
            )}
    
            {view === 'summary' && (
                <div className="bg-white p-6 rounded-xl shadow-lg">
                    <div className="flex flex-col md:flex-row justify-between items-center mb-6 gap-4 no-print">
                        <h2 className="text-2xl font-bold text-gray-800">ملخص إجازات الموظفين</h2>
                         <div className="flex items-center gap-2 flex-wrap justify-center">
                            <button onClick={handleCopyToClipboard} className="bg-gray-700 hover:bg-gray-800 text-white font-bold py-2 px-4 rounded-lg flex items-center gap-2 transition-colors">
                                <ClipboardDocumentIcon className="w-5 h-5"/>
                                <span>نسخ كجدول</span>
                            </button>
                             <button onClick={handleExportSummaryToWord} className="bg-blue-600 hover:bg-blue-700 text-white font-bold py-2 px-4 rounded-lg flex items-center gap-2 transition-colors">
                                <DocumentTextIcon className="w-5 h-5"/>
                                <span>تصدير Word</span>
                            </button>
                             <button onClick={handlePrint} className="bg-sky-500 hover:bg-sky-600 text-white font-bold py-2 px-4 rounded-lg flex items-center gap-2 transition-colors">
                                <PrinterIcon className="w-5 h-5"/>
                                <span>طباعة</span>
                            </button>
                             <button onClick={handleExportPDF} className="bg-red-500 hover:bg-red-600 text-white font-bold py-2 px-4 rounded-lg flex items-center gap-2 transition-colors">
                                <PdfIcon className="w-5 h-5"/>
                                <span>تصدير PDF</span>
                            </button>
                         </div>
                    </div>
                     {currentUser === 'admin' && (
                         <div className="grid grid-cols-1 md:grid-cols-3 gap-4 mb-6 no-print">
                            <div className="relative">
                                <input
                                    type="text"
                                    placeholder="...ابحث عن موظف"
                                    value={summarySearchTerm}
                                    onChange={(e) => setSummarySearchTerm(e.target.value)}
                                    className="w-full pl-10 pr-4 py-2 border rounded-lg focus:outline-none focus:ring-2 focus:ring-indigo-500"
                                />
                                <SearchIcon className="absolute right-3 top-1/2 transform -translate-y-1/2 w-5 h-5 text-gray-400" />
                            </div>
                            <select
                                value={selectedYear}
                                onChange={(e) => setSelectedYear(e.target.value)}
                                className="w-full p-2 border rounded-lg focus:outline-none focus:ring-2 focus:ring-indigo-500"
                            >
                                <option value="all">كل السنوات</option>
                                {availableYears.map(year => <option key={year} value={year}>{year}</option>)}
                            </select>
                             <select
                                value={selectedMonth}
                                onChange={(e) => setSelectedMonth(e.target.value)}
                                className="w-full p-2 border rounded-lg focus:outline-none focus:ring-2 focus:ring-indigo-500"
                                disabled={selectedYear === 'all'}
                            >
                                <option value="all">كل الأشهر</option>
                                {Array.from({ length: 12 }, (_, i) => i + 1).map(month => (
                                    <option key={month} value={month}>{new Date(2000, month-1, 1).toLocaleString('ar', { month: 'long' })}</option>
                                ))}
                            </select>
                        </div>
                    )}
                    
                    <div ref={summaryContentRef} className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 xl:grid-cols-4 gap-6">
                        {displayedSummary.map((employee, index) => (
                            <EmployeeSummaryCard key={`${employee.name}-${index}`} employee={employee} />
                        ))}
                    </div>
                    {displayedSummary.length === 0 && <p className="text-center text-gray-500 mt-8">.لا توجد بيانات لعرضها</p>}
                </div>
            )}
            
            {view === 'rankedSummary' && (
                <div className="bg-white p-6 rounded-xl shadow-lg">
                    <div className="flex flex-col md:flex-row justify-between items-center mb-6 gap-4 no-print">
                        <h2 className="text-2xl font-bold text-gray-800">الملخص المرتب</h2>
                         <div className="flex items-center gap-4">
                            <span className="font-semibold">طريقة الترتيب:</span>
                            <div className="flex items-center gap-2 rounded-lg p-1 bg-gray-100">
                                 <button onClick={() => setRankedSortOrder('byDays')} className={`px-3 py-1 text-sm rounded-md flex items-center gap-2 ${rankedSortOrder === 'byDays' ? 'bg-indigo-600 text-white shadow' : 'hover:bg-gray-200'}`}>
                                    <BarsArrowUpIcon className="w-5 h-5" />
                                    <span>حسب الأيام</span>
                                 </button>
                                 <button onClick={() => setRankedSortOrder('alphabetical')} className={`px-3 py-1 text-sm rounded-md flex items-center gap-2 ${rankedSortOrder === 'alphabetical' ? 'bg-indigo-600 text-white shadow' : 'hover:bg-gray-200'}`}>
                                    <AlphabeticalSortIcon className="w-5 h-5" />
                                    <span>أبجدي</span>
                                 </button>
                            </div>
                         </div>
                         <div className="flex items-center gap-2">
                            <button onClick={handleExportWord} className="bg-blue-600 hover:bg-blue-700 text-white font-bold py-2 px-4 rounded-lg flex items-center gap-2 transition-colors">
                                <DocumentTextIcon className="w-5 h-5"/>
                                <span>تصدير Word</span>
                            </button>
                             <button onClick={handlePrint} className="bg-sky-500 hover:bg-sky-600 text-white font-bold py-2 px-4 rounded-lg flex items-center gap-2 transition-colors">
                                <PrinterIcon className="w-5 h-5"/>
                                <span>طباعة</span>
                            </button>
                         </div>
                    </div>
    
                    <div ref={summaryContentRef}>
                        {renderRankedSummary()}
                    </div>
                </div>
            )}

            {view === 'monthlyReport' && (
                <div className="bg-white p-6 rounded-xl shadow-lg">
                    <div className="flex flex-col md:flex-row justify-between items-center mb-6 gap-4 no-print">
                         <h2 className="text-2xl font-bold text-gray-800">التقرير الشهري النهائي</h2>
                          {monthlyReportData.length > 0 && (
                            <button onClick={handleExportMonthlyReportToWord} className="bg-blue-600 hover:bg-blue-700 text-white font-bold py-2 px-4 rounded-lg flex items-center gap-2 transition-colors">
                                <DocumentTextIcon className="w-5 h-5"/>
                                <span>تصدير Word</span>
                            </button>
                          )}
                    </div>
                    <div className="grid grid-cols-1 md:grid-cols-3 gap-4 mb-6 no-print p-4 border rounded-lg bg-gray-50">
                        <select
                            value={reportYear}
                            onChange={(e) => setReportYear(Number(e.target.value))}
                            className="w-full p-2 border rounded-lg focus:outline-none focus:ring-2 focus:ring-indigo-500"
                        >
                            {Array.from({ length: 10 }, (_, i) => new Date().getFullYear() - i).map(year => 
                                <option key={year} value={year}>{toArabicNumerals(year)}</option>
                            )}
                        </select>
                        <select
                            value={reportMonth}
                            onChange={(e) => setReportMonth(Number(e.target.value))}
                            className="w-full p-2 border rounded-lg focus:outline-none focus:ring-2 focus:ring-indigo-500"
                        >
                            {Array.from({ length: 12 }, (_, i) => i + 1).map(month => (
                                <option key={month} value={month}>{new Date(2000, month-1, 1).toLocaleString('ar', { month: 'long' })}</option>
                            ))}
                        </select>
                        <button onClick={handleGenerateMonthlyReport} className="bg-indigo-600 hover:bg-indigo-700 text-white font-bold py-2 px-4 rounded-lg flex items-center justify-center gap-2 transition-colors">
                            <CalendarDaysIcon className="w-5 h-5"/>
                            <span>إنشاء التقرير</span>
                        </button>
                    </div>

                    {monthlyReportData.length > 0 && (
                        <div className="overflow-x-auto">
                            <table className="w-full text-sm text-right text-gray-500 border">
                                <thead className="text-xs text-gray-700 uppercase bg-gray-100">
                                    <tr>
                                        <th className="px-2 py-3 border">ت</th>
                                        <th className="px-4 py-3 border">اسم الموظف</th>
                                        <th className="px-2 py-3 border">الرصيد الأولي</th>
                                        <th className="px-4 py-3 border">الاعتيادية وتاريخها</th>
                                        <th className="px-4 py-3 border">الزمنية (أيام/ساعات)</th>
                                        <th className="px-4 py-3 border">المرضية</th>
                                        <th className="px-4 py-3 border">الطويلة</th>
                                        <th className="px-2 py-3 border">الرصيد المتبقي</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    {monthlyReportData.map((row, index) => (
                                        <tr key={index} className="bg-white border-b hover:bg-gray-50 text-center">
                                            <td className="px-2 py-2 border">{toArabicNumerals(index + 1)}</td>
                                            <td className="px-4 py-2 border text-right font-medium text-gray-800">{row.name}</td>
                                            <td className="px-2 py-2 border">{toArabicNumerals(row.initialBalance)}</td>
                                            <td className="px-4 py-2 border">
                                                {row.regularLeaves.count > 0 ? `${formatDaysArabic(row.regularLeaves.count)} /${row.regularLeaves.dates}` : '-'}
                                            </td>
                                            <td className="px-4 py-2 border">
                                                {(row.hourlyLeaves.days > 0 || row.hourlyLeaves.hours > 0) ? `${toArabicNumerals(row.hourlyLeaves.days)} يوم / ${toArabicNumerals(parseFloat(row.hourlyLeaves.hours.toFixed(2)))} س` : '-'}
                                            </td>
                                            <td className="px-4 py-2 border">{row.sickLeave.dateRange || '-'}</td>
                                            <td className="px-4 py-2 border">{row.longLeave.dateRange || '-'}</td>
                                            <td className="px-2 py-2 border font-bold">{toArabicNumerals(row.finalBalance)}</td>
                                        </tr>
                                    ))}
                                </tbody>
                            </table>
                        </div>
                    )}
                </div>
            )}
            
            {view === 'employeeManagement' && (
                <div className="space-y-8">
                    {/* Add New Employee Form */}
                    <div className="bg-white p-6 rounded-xl shadow-lg">
                        <h2 className="text-2xl font-bold text-gray-800 mb-4">إضافة موظف جديد</h2>
                        <form onSubmit={handleAddEmployee} className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-6 items-end">
                            <div className="lg:col-span-1">
                                <label className="block text-sm font-medium text-gray-700">الاسم الكامل</label>
                                <input type="text" value={newEmployeeName} onChange={e => setNewEmployeeName(e.target.value)} required className="mt-1 block w-full px-3 py-2 bg-white border border-gray-300 rounded-md shadow-sm focus:outline-none focus:ring-indigo-500 focus:border-indigo-500"/>
                            </div>
                            <div>
                                <label className="block text-sm font-medium text-gray-700">الرصيد الأولي</label>
                                <input type="number" value={newEmployeeBalance} onChange={e => setNewEmployeeBalance(e.target.value)} required className="mt-1 block w-full px-3 py-2 bg-white border border-gray-300 rounded-md shadow-sm focus:outline-none focus:ring-indigo-500 focus:border-indigo-500"/>
                            </div>
                             <div>
                                <label className="block text-sm font-medium text-gray-700">ساعات العمل باليوم</label>
                                <input type="number" value={newEmployeeWorkdayHours} onChange={e => setNewEmployeeWorkdayHours(e.target.value)} required className="mt-1 block w-full px-3 py-2 bg-white border border-gray-300 rounded-md shadow-sm focus:outline-none focus:ring-indigo-500 focus:border-indigo-500"/>
                            </div>
                            <div className="flex gap-2">
                                <div>
                                    <label className="block text-sm font-medium text-gray-700">اسم المستخدم</label>
                                    <input type="text" value={newEmployeeUsername} onChange={e => setNewEmployeeUsername(e.target.value)} required className="mt-1 block w-full px-3 py-2 bg-white border border-gray-300 rounded-md shadow-sm focus:outline-none focus:ring-indigo-500 focus:border-indigo-500"/>
                                </div>
                                <button type="button" onClick={() => setNewEmployeeUsername(generateRandomUsername())} className="mt-6 p-2 bg-gray-200 rounded-md hover:bg-gray-300 self-end">
                                    <ArrowPathIcon className="w-5 h-5"/>
                                </button>
                            </div>
                             <div className="flex gap-2">
                                <div>
                                    <label className="block text-sm font-medium text-gray-700">كلمة المرور</label>
                                    <input type="text" value={newEmployeePassword} onChange={e => setNewEmployeePassword(e.target.value)} required className="mt-1 block w-full px-3 py-2 bg-white border border-gray-300 rounded-md shadow-sm focus:outline-none focus:ring-indigo-500 focus:border-indigo-500"/>
                                </div>
                                 <button type="button" onClick={() => setNewEmployeePassword(generateRandomPassword())} className="mt-6 p-2 bg-gray-200 rounded-md hover:bg-gray-300 self-end">
                                    <ArrowPathIcon className="w-5 h-5"/>
                                </button>
                            </div>
                             <div className="md:col-span-2 lg:col-span-2">
                                <label className="block text-sm font-medium text-gray-700">الصورة الشخصية (اختياري)</label>
                                <div className="mt-1 flex items-center gap-4">
                                    {newEmployeePhoto ? (
                                        <img src={newEmployeePhoto} alt="Preview" className="w-12 h-12 rounded-full object-cover"/>
                                    ) : (
                                        <div className="w-12 h-12 rounded-full bg-gray-200 flex items-center justify-center">
                                            <UserCircleIcon className="w-8 h-8 text-gray-400"/>
                                        </div>
                                    )}
                                    <input type="file" accept="image/*" onChange={handleNewEmployeePhotoChange} className="block w-full text-sm text-slate-500 file:mr-4 file:py-2 file:px-4 file:rounded-full file:border-0 file:text-sm file:font-semibold file:bg-violet-50 file:text-violet-700 hover:file:bg-violet-100"/>
                                </div>
                            </div>
                            <button type="submit" className="bg-indigo-600 hover:bg-indigo-700 text-white font-bold py-2 px-4 rounded-lg transition-colors h-10">إضافة موظف</button>
                        </form>
                    </div>
    
                    {/* Employee List */}
                    <div className="bg-white p-6 rounded-xl shadow-lg">
                        <div className="flex flex-col md:flex-row justify-between items-center mb-4 gap-4">
                             <h2 className="text-2xl font-bold text-gray-800">قائمة الموظفين ({toArabicNumerals(employees.length)})</h2>
                            <div className="flex items-center gap-2 flex-wrap justify-center md:justify-end">
                                <div className="relative">
                                    <input type="text" placeholder="...ابحث" value={employeeSearchTerm} onChange={e => setEmployeeSearchTerm(e.target.value)} className="w-full pl-10 pr-4 py-2 border rounded-lg focus:outline-none focus:ring-2 focus:ring-indigo-500" />
                                    <SearchIcon className="absolute right-3 top-1/2 transform -translate-y-1/2 w-5 h-5 text-gray-400" />
                                </div>
                                <button onClick={handleCorrectBalances} className="bg-orange-500 hover:bg-orange-600 text-white font-bold py-2 px-4 rounded-lg flex items-center gap-2 transition-colors">
                                    <MinusCircleIcon className="w-5 h-5"/>
                                    <span>تصحيح الرصيد (-5)</span>
                                </button>
                                <button onClick={handleExportUserList} className="bg-green-600 hover:bg-green-700 text-white font-bold py-2 px-4 rounded-lg flex items-center gap-2 transition-colors">
                                    <DownloadIcon className="w-5 h-5"/>
                                    <span>تصدير القائمة</span>
                                </button>
                            </div>
                        </div>
                        
                        <div className="overflow-x-auto">
                            <table className="w-full text-sm text-right text-gray-500">
                                <thead className="text-xs text-gray-700 uppercase bg-gray-100">
                                    <tr>
                                        <th scope="col" className="px-4 py-3">الصورة</th>
                                        <th scope="col" className="px-6 py-3">الاسم</th>
                                        <th scope="col" className="px-6 py-3">الرصيد الحالي</th>
                                        <th scope="col" className="px-6 py-3">ساعات العمل</th>
                                        <th scope="col" className="px-6 py-3">اسم المستخدم</th>
                                        <th scope="col" className="px-6 py-3">كلمة المرور</th>
                                        <th scope="col" className="px-6 py-3 text-center">الإجراءات</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    {filteredEmployees.map(emp => (
                                        <tr key={emp.id} className="bg-white border-b hover:bg-gray-50">
                                            {editingEmployee?.id === emp.id ? (
                                                <>
                                                    <td className="px-4 py-2">
                                                        <div className="flex flex-col items-center gap-2">
                                                            {editingEmployee.photo ? (
                                                                <img src={editingEmployee.photo} alt={editingEmployee.name} className="w-10 h-10 rounded-full object-cover" />
                                                            ) : (
                                                                <div className="w-10 h-10 rounded-full bg-gray-200 flex items-center justify-center"><UserCircleIcon className="w-6 h-6 text-gray-400" /></div>
                                                            )}
                                                            <input type="file" accept="image/*" onChange={handleEditPhotoChange} className="w-24 text-xs"/>
                                                            <button onClick={() => setEditingEmployee({...editingEmployee, photo: undefined})} className="text-xs text-red-500 hover:text-red-700">إزالة</button>
                                                        </div>
                                                    </td>
                                                    <td className="px-6 py-2 align-middle">
                                                        <input type="text" value={editingEmployee.name} onChange={e => setEditingEmployee({...editingEmployee, name: e.target.value})} className="w-full p-1 border rounded" />
                                                    </td>
                                                    <td className="px-6 py-2 align-middle">
                                                        <input type="number" value={editingEmployee.balance} onChange={e => setEditingEmployee({...editingEmployee, balance: Number(e.target.value)})} className="w-20 p-1 border rounded" />
                                                    </td>
                                                    <td className="px-6 py-2 align-middle">
                                                        <input type="number" value={editingEmployee.workdayHours} onChange={e => setEditingEmployee({...editingEmployee, workdayHours: Number(e.target.value)})} className="w-20 p-1 border rounded" />
                                                    </td>
                                                    <td className="px-6 py-2 align-middle">
                                                        <input type="text" value={editingEmployee.username} onChange={e => setEditingEmployee({...editingEmployee, username: e.target.value})} className="w-24 p-1 border rounded" />
                                                    </td>
                                                    <td className="px-6 py-2 align-middle">
                                                        <input type="text" value={editingEmployee.password} onChange={e => setEditingEmployee({...editingEmployee, password: e.target.value})} className="w-24 p-1 border rounded" />
                                                    </td>
                                                    <td className="px-6 py-2 text-center align-middle">
                                                        <div className="flex justify-center gap-2">
                                                            <button onClick={handleSaveEdit} className="text-green-600 hover:text-green-800 font-semibold disabled:text-gray-400 disabled:cursor-not-allowed" disabled={isImageProcessing}>
                                                                {isImageProcessing ? '...جاري' : 'حفظ'}
                                                            </button>
                                                            <button onClick={handleCancelEdit} className="text-gray-500 hover:text-gray-700">إلغاء</button>
                                                        </div>
                                                    </td>
                                                </>
                                            ) : (
                                                <>
                                                    <td className="px-4 py-2">
                                                        {emp.photo ? (
                                                             <img src={emp.photo} alt={emp.name} className="w-10 h-10 rounded-full object-cover" />
                                                        ) : (
                                                            <div className="w-10 h-10 rounded-full bg-gray-200 flex items-center justify-center"><UserCircleIcon className="w-6 h-6 text-gray-400" /></div>
                                                        )}
                                                    </td>
                                                    <th scope="row" className="px-6 py-4 font-medium text-gray-900 whitespace-nowrap">{emp.name}</th>
                                                    <td className="px-6 py-4">{toArabicNumerals(emp.balance)}</td>
                                                    <td className="px-6 py-4">{toArabicNumerals(emp.workdayHours)}</td>
                                                    <td className="px-6 py-4">{emp.username}</td>
                                                    <td className="px-6 py-4">{emp.password}</td>
                                                    <td className="px-6 py-4 text-center">
                                                        <div className="flex justify-center gap-4">
                                                            <button onClick={() => handleStartEdit(emp)} className="text-indigo-600 hover:text-indigo-800">
                                                                <PencilIcon className="w-5 h-5" />
                                                            </button>
                                                            <button onClick={() => handleDeleteEmployee(emp.id)} className="text-red-600 hover:text-red-800">
                                                                 <TrashIcon className="w-5 h-5" />
                                                            </button>
                                                        </div>
                                                    </td>
                                                </>
                                            )}
                                        </tr>
                                    ))}
                                </tbody>
                            </table>
                        </div>
                    </div>
                </div>
            )}
            </div>
          </main>
        </div>
      );
};