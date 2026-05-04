const express = require('express');
const multer = require('multer');
const cors = require('cors');
const path = require('path');
const fs = require('fs');
const crypto = require('crypto'); // For SHA-256 hashing
const { promisify } = require('util');
const sqlite3 = require('sqlite3').verbose();
const OpenAI = require('openai');
const Anthropic = require('@anthropic-ai/sdk');
const { DocumentProcessorServiceClient } = require('@google-cloud/documentai');
const axios = require('axios');
const FormData = require('form-data');
const XLSX = require('xlsx'); // Added for Tally Excel export
const dayjs = require('dayjs');
const customParseFormat = require('dayjs/plugin/customParseFormat');
const utc = require('dayjs/plugin/utc');
const timezone = require('dayjs/plugin/timezone');
const advancedFormat = require('dayjs/plugin/advancedFormat');
const localizedFormat = require('dayjs/plugin/localizedFormat');
const { XMLBuilder } = require('fast-xml-parser'); // Added for Tally XML generation
const { initTallySchema } = require('./tally/db-schema');
const tallyRoutes = require('./tally/tally-routes');
const tallyPushRoutes = require('./tally/tally-push');

//Naman
// Enable plugins
dayjs.extend(customParseFormat);
dayjs.extend(utc);
dayjs.extend(timezone);
dayjs.extend(advancedFormat);
dayjs.extend(localizedFormat);

// Load environment variables (.env always wins over pre-existing empty system env vars)
require('dotenv').config({ override: true });

let sharp;
let pdfToImg;
let imageOptimizationAvailable = false;
let pdfConversionAvailable = false;

// canvas must be required BEFORE sharp on Windows to avoid DLL conflicts
try { require('canvas'); } catch (_) {}

// Load optional dependencies
try {
    sharp = require('sharp');
    imageOptimizationAvailable = true;
    console.log('✅ Image processing available');
} catch (error) {
    console.log('⚠️  Basic image processing mode');
    imageOptimizationAvailable = false;
}

// Initialize PDF converter
const initializePdfConverter = async () => {
    try {
        const pdfModule = await import('pdf-to-img');
        pdfToImg = pdfModule.pdf;
        if (typeof pdfToImg !== 'function') {
            throw new Error(`pdf-to-img exported unexpected type: ${typeof pdfToImg}`);
        }
        pdfConversionAvailable = true;
        console.log('✅ PDF conversion available');
    } catch (error) {
        console.log('⚠️  PDF conversion not available:', error.message);
        pdfConversionAvailable = false;
    }
};

const app = express();
const PORT = 3001;

app.use(cors());
app.use(express.json());
app.use(express.static('public'));

// Tally integration routes
app.use('/api', tallyRoutes);
app.use('/api', tallyPushRoutes);

// 🆕 NEW LEDGER STANDARDIZATION CLASS
class LedgerStandardizer {
    constructor() {
        this.ledgerAliasMap = new Map();
        this.initialized = false;
    }

    /**
     * Load ledger master file and build alias mapping
     * @returns {Promise<boolean>} - Success status
     */
    async loadLedgerMaster() {
        try {
            const ledgerMasterPath = './Ledger extractionV2/TallyData_Complete_latest.xlsx';
            
            console.log('📚 Loading ledger master file for standardization:', ledgerMasterPath);
            
            // Check if file exists
            if (!fs.existsSync(ledgerMasterPath)) {
                console.warn('⚠️ Ledger master file not found:', ledgerMasterPath);
                console.warn('   Ledger standardization will be skipped');
                this.initialized = false;
                return false;
            }

            // Read the Excel file
            const workbook = XLSX.readFile(ledgerMasterPath);
            
            // Check if the '📋 Ledgers' sheet exists
            const sheetName = '📋 Ledgers';
            if (!workbook.SheetNames.includes(sheetName)) {
                console.warn(`⚠️ Sheet '${sheetName}' not found in ledger master file`);
                console.warn('   Available sheets:', workbook.SheetNames);
                this.initialized = false;
                return false;
            }

            // Convert sheet to JSON
            const worksheet = workbook.Sheets[sheetName];
            const ledgerData = XLSX.utils.sheet_to_json(worksheet);
            
            console.log(`📊 Found ${ledgerData.length} ledger records in master file`);

            // Clear existing mapping
            this.ledgerAliasMap.clear();

            // Build the alias mapping
            let aliasCount = 0;
            
            ledgerData.forEach((row, index) => {
                const officialName = row['Ledger Name'];
                const aliases = row['Aliases'];
                
                if (!officialName || officialName.trim().length === 0) {
                    console.warn(`⚠️ Row ${index + 1}: Missing ledger name, skipping`);
                    return;
                }

                const trimmedOfficialName = officialName.trim();
                
                // Add the official name itself as a key (self-mapping)
                const officialKey = trimmedOfficialName.toLowerCase();
                this.ledgerAliasMap.set(officialKey, trimmedOfficialName);
                aliasCount++;
                
                // Process aliases if they exist
                if (aliases && aliases.trim().length > 0) {
                    const aliasArray = aliases.split(',').map(alias => alias.trim()).filter(alias => alias.length > 0);
                    
                    aliasArray.forEach(alias => {
                        const aliasKey = alias.toLowerCase();
                        if (!this.ledgerAliasMap.has(aliasKey)) {
                            this.ledgerAliasMap.set(aliasKey, trimmedOfficialName);
                            aliasCount++;
                        } else {
                            console.warn(`⚠️ Duplicate alias '${alias}' found for '${trimmedOfficialName}', keeping existing mapping to '${this.ledgerAliasMap.get(aliasKey)}'`);
                        }
                    });
                }
            });

            this.initialized = true;
            console.log(`✅ Ledger standardization initialized: ${aliasCount} aliases mapped to ${ledgerData.length} official ledger names`);
            
            // Log some sample mappings for verification
            if (this.ledgerAliasMap.size > 0) {
                console.log('📝 Sample alias mappings:');
                let sampleCount = 0;
                for (const [alias, official] of this.ledgerAliasMap) {
                    if (sampleCount < 5) {
                        console.log(`   "${alias}" → "${official}"`);
                        sampleCount++;
                    } else {
                        break;
                    }
                }
                if (this.ledgerAliasMap.size > 5) {
                    console.log(`   ... and ${this.ledgerAliasMap.size - 5} more mappings`);
                }
            }

            return true;

        } catch (error) {
            console.error('❌ Failed to load ledger master file:', error.message);
            this.initialized = false;
            return false;
        }
    }

    /**
     * Standardize a ledger name using the alias mapping
     * @param {string} ledgerName - Original ledger name
     * @returns {string} - Standardized ledger name or original if no match
     */
    standardizeLedgerName(ledgerName) {
        if (!this.initialized || !ledgerName || typeof ledgerName !== 'string') {
            return ledgerName;
        }

        const trimmedName = ledgerName.trim();
        const lookupKey = trimmedName.toLowerCase();
        
        if (this.ledgerAliasMap.has(lookupKey)) {
            const standardizedName = this.ledgerAliasMap.get(lookupKey);
            if (standardizedName !== trimmedName) {
                console.log(`🔄 Ledger standardized: "${trimmedName}" → "${standardizedName}"`);
            }
            return standardizedName;
        }

        // No match found, return original
        return trimmedName;
    }

    /**
     * Standardize ledger names in Tally export data
     * @param {Array} tallyData - Tally export data array
     * @returns {Array} - Data with standardized ledger names
     */
    standardizeTallyExportData(tallyData) {
        if (!this.initialized || !Array.isArray(tallyData)) {
            console.log('⚠️ Ledger standardization skipped: not initialized or invalid data');
            return tallyData;
        }

        console.log(`🔄 Standardizing ledger names in ${tallyData.length} export entries...`);
        
        let standardizedCount = 0;
        
        tallyData.forEach(entry => {
            if (entry['Ledger Name']) {
                const originalName = entry['Ledger Name'];
                const standardizedName = this.standardizeLedgerName(originalName);
                
                if (standardizedName !== originalName) {
                    entry['Ledger Name'] = standardizedName;
                    standardizedCount++;
                }
            }
        });

        console.log(`✅ Ledger standardization complete: ${standardizedCount} entries updated`);
        return tallyData;
    }

    /**
     * Get statistics about the loaded alias mapping
     * @returns {object} - Statistics object
     */
    getStatistics() {
        return {
            initialized: this.initialized,
            totalAliases: this.ledgerAliasMap.size,
            sampleMappings: Array.from(this.ledgerAliasMap.entries()).slice(0, 5)
        };
    }
}

// 🔐 ENHANCED DUPLICATE DETECTION UTILITIES WITH CONTENT-BASED HASHING
class DuplicateDetector {
    /**
     * Computes SHA-256 hash of a file using streaming for memory efficiency
     * @param {string} filePath - Path to the file
     * @returns {Promise<string>} - SHA-256 hash in hexadecimal
     */
    static async computeFileHash(filePath) {
        return new Promise((resolve, reject) => {
            const hash = crypto.createHash('sha256');
            const stream = fs.createReadStream(filePath);
            
            stream.on('data', (data) => {
                hash.update(data);
            });
            
            stream.on('end', () => {
                resolve(hash.digest('hex'));
            });
            
            stream.on('error', (error) => {
                reject(error);
            });
        });
    }

    /**
     * Checks if a file with the same content hash already exists in the database
     * @param {string} fileHash - SHA-256 hash of the file
     * @returns {Promise<object|null>} - Existing file record or null
     */
    static async checkFileHashExists(fileHash) {
        return new Promise((resolve, reject) => {
            const sql = `SELECT id, original_filename, file_hash, file_path, invoice_number, vendor_name, 
                        parseur_document_id, created_at, confidence_score, processing_source 
                        FROM expenses 
                        WHERE file_hash = ? 
                        ORDER BY created_at DESC 
                        LIMIT 1`;
            
            db.get(sql, [fileHash], (err, row) => {
                if (err) {
                    reject(err);
                } else {
                    resolve(row);
                }
            });
        });
    }

    /**
     * Checks if a file with the same name already exists in the database
     * @param {string} originalName - Original filename
     * @returns {Promise<object|null>} - Existing file record or null
     */
    static async checkFileNameExists(originalName) {
        return new Promise((resolve, reject) => {
            const sql = `SELECT id, original_filename, file_hash, file_path, invoice_number, vendor_name, 
                        parseur_document_id, created_at 
                        FROM expenses 
                        WHERE original_filename = ? 
                        ORDER BY created_at DESC 
                        LIMIT 1`;
            
            db.get(sql, [originalName], (err, row) => {
                if (err) {
                    reject(err);
                } else {
                    resolve(row);
                }
            });
        });
    }

    /**
     * Safely deletes a file if it exists
     * @param {string} filePath - Path to the file to delete
     */
    static async safeDeleteFile(filePath) {
        try {
            if (fs.existsSync(filePath)) {
                await fs.promises.unlink(filePath);
                console.log(`🗑️  Deleted duplicate file: ${filePath}`);
            }
        } catch (error) {
            console.error(`❌ Failed to delete file ${filePath}:`, error.message);
        }
    }

    /**
     * Enhanced duplicate detection logic with content-based hashing
     * @param {object} uploadedFile - Multer file object
     * @returns {Promise<object>} - Validation result with reprocessing options
     */
    static async validateUpload(uploadedFile) {
        try {
            // Step 1: Compute file hash for content-based duplicate detection
            console.log('🔍 Computing SHA-256 hash for content-based duplicate detection...');
            const newFileHash = await this.computeFileHash(uploadedFile.path);
            
            // Step 2: Check if content already exists (primary duplicate detection)
            const existingFileByHash = await this.checkFileHashExists(newFileHash);
            
            if (existingFileByHash) {
                // Content duplicate found - delete uploaded file and offer reprocessing
                await this.safeDeleteFile(uploadedFile.path);
                
                const canReprocess = !!(existingFileByHash.parseur_document_id);
                
                return {
                    isValid: false,
                    isDuplicate: true,
                    duplicateType: 'content',
                    message: 'Duplicate file content detected',
                    details: `A file with identical content has already been processed`,
                    fileHash: newFileHash,
                    existingFile: {
                        id: existingFileByHash.id,
                        original_filename: existingFileByHash.original_filename,
                        invoice_number: existingFileByHash.invoice_number,
                        vendor_name: existingFileByHash.vendor_name,
                        parseur_document_id: existingFileByHash.parseur_document_id,
                        uploaded_at: existingFileByHash.created_at,
                        confidence_score: existingFileByHash.confidence_score,
                        processing_source: existingFileByHash.processing_source
                    },
                    reprocessAllowed: canReprocess,
                    reprocessMessage: canReprocess ? 
                        'You can reprocess this document using the existing Parseur data' : 
                        'Reprocessing not available - no Parseur document ID found'
                };
            }

            // Step 3: Check filename conflicts (secondary check)
            const existingFileByName = await this.checkFileNameExists(uploadedFile.originalname);
            
            if (existingFileByName) {
                // Same filename but different content - potential issue
                console.log('⚠️ Filename conflict detected - same name, different content');
                
                return {
                    isValid: true, // Allow upload but warn user
                    isNameConflict: true,
                    message: 'Filename conflict detected',
                    details: `A different file with the name "${uploadedFile.originalname}" exists. The content is different, so upload will proceed.`,
                    fileHash: newFileHash,
                    existingFile: {
                        id: existingFileByName.id,
                        original_filename: existingFileByName.original_filename,
                        invoice_number: existingFileByName.invoice_number,
                        vendor_name: existingFileByName.vendor_name,
                        uploaded_at: existingFileByName.created_at
                    },
                    warning: 'Consider renaming the file to avoid confusion'
                };
            }

            // Step 4: File is unique - proceed with processing
            return {
                isValid: true,
                message: 'File validated successfully - unique content and filename',
                fileHash: newFileHash,
                duplicateCheck: 'passed'
            };

        } catch (error) {
            // On error, clean up the uploaded file
            await this.safeDeleteFile(uploadedFile.path);
            console.error('❌ Duplicate detection failed:', error);
            throw new Error(`Duplicate detection failed: ${error.message}`);
        }
    }

    /**
     * Enhanced validation for bulk uploads with batch optimization
     * @param {Array} uploadedFiles - Array of multer file objects
     * @returns {Promise<object>} - Batch validation results
     */
    static async validateBulkUpload(uploadedFiles) {
        const results = {
            valid: [],
            duplicates: [],
            conflicts: [],
            errors: [],
            batchStats: {
                total: uploadedFiles.length,
                unique: 0,
                contentDuplicates: 0,
                nameConflicts: 0,
                errors: 0
            }
        };

        // Compute hashes for all files first
        const fileHashes = new Map();
        for (const file of uploadedFiles) {
            try {
                const hash = await this.computeFileHash(file.path);
                fileHashes.set(file.originalname, hash);
            } catch (error) {
                results.errors.push({
                    filename: file.originalname,
                    error: `Hash computation failed: ${error.message}`
                });
                await this.safeDeleteFile(file.path);
            }
        }

        // Track which hashes have already been seen in this batch
        // so we let the FIRST occurrence through and reject subsequent copies
        const seenHashesInBatch = new Set();

        // Process each file
        for (const file of uploadedFiles) {
            if (!fileHashes.has(file.originalname)) continue; // Skip files with hash errors

            const fileHash = fileHashes.get(file.originalname);

            // Check for batch-internal duplicates — only reject 2nd+ copies, let the first through
            if (seenHashesInBatch.has(fileHash)) {
                results.duplicates.push({
                    filename: file.originalname,
                    error: 'Same file uploaded twice in this batch',
                    type: 'batch_internal'
                });
                results.batchStats.contentDuplicates++;
                await this.safeDeleteFile(file.path);
                continue;
            }
            seenHashesInBatch.add(fileHash);

            // Check against database
            try {
                const validation = await this.validateUpload(file);
                
                if (validation.isValid) {
                    if (validation.isNameConflict) {
                        results.conflicts.push({
                            filename: file.originalname,
                            validation: validation
                        });
                        results.batchStats.nameConflicts++;
                    } else {
                        results.valid.push({
                            filename: file.originalname,
                            validation: validation
                        });
                        results.batchStats.unique++;
                    }
                } else if (validation.isDuplicate) {
                    results.duplicates.push({
                        filename: file.originalname,
                        error: validation.message || 'File already uploaded previously',
                        type: 'db_duplicate',
                        validation: validation
                    });
                    results.batchStats.contentDuplicates++;
                }
            } catch (error) {
                results.errors.push({
                    filename: file.originalname,
                    error: error.message
                });
                results.batchStats.errors++;
            }
        }

        return results;
    }
}

// 🆕 ENHANCED UNIVERSAL DATE PARSER CLASS
class UniversalDateParser {
    static parseInvoiceDate(dateStr) {
        if (!dateStr || typeof dateStr !== 'string') {
            console.warn('⚠️ Invalid date input:', dateStr);
            return '';
        }

        // Clean the input
        const cleanDate = dateStr.toString().trim()
            .replace(/[,\u00A0]/g, '') // Remove commas and non-breaking spaces
            .replace(/\s+/g, ' ')      // Normalize spaces
            .replace(/^(dated|date|dt)[\s:]+/i, '') // Remove prefixes like "Date:", "Dated:"
            .trim();

        console.log('🔍 Parsing date:', dateStr, '→ cleaned:', cleanDate);

        // Define common date formats found in invoices
        const dateFormats = [
            // ISO formats
            'YYYY-MM-DD',
            'YYYY/MM/DD',
            
            // European formats
            'DD-MM-YYYY',
            'DD/MM/YYYY',
            'DD.MM.YYYY',
            'DD-MM-YY',
            'DD/MM/YY',
            'DD.MM.YY',
            
            // US formats
            'MM-DD-YYYY',
            'MM/DD/YYYY',
            'MM.DD.YYYY',
            'MM-DD-YY',
            'MM/DD/YY',
            
            // Month name formats (your example: "21 January, 2019")
            'DD MMMM YYYY',
            'DD MMM YYYY',
            'DD MMMM, YYYY',
            'DD MMM, YYYY',
            'D MMMM YYYY',
            'D MMM YYYY',
            'D MMMM, YYYY',
            'D MMM, YYYY',
            
            // Month first formats
            'MMMM DD, YYYY',
            'MMM DD, YYYY',
            'MMMM DD YYYY',
            'MMM DD YYYY',
            'MMMM D, YYYY',
            'MMM D, YYYY',
            'MMMM D YYYY',
            'MMM D YYYY',
            
            // Year first with month names
            'YYYY MMMM DD',
            'YYYY MMM DD',
            'YYYY MMMM D',
            'YYYY MMM D',
            
            // Indian formats
            'DD-MMM-YYYY',
            'DD-MMM-YY',
            'DD/MMM/YYYY',
            'DD/MMM/YY',
            
            // Compact formats
            'DDMMYYYY',
            'DDMMYY',
            'YYYYMMDD',
            'YYMMDD',
            
            // Special formats with ordinals
            'MMMM Do, YYYY',
            'MMM Do, YYYY',
            'Do MMMM YYYY',
            'Do MMM YYYY',
            
            // Time included formats (will extract date part)
            'YYYY-MM-DD HH:mm:ss',
            'DD/MM/YYYY HH:mm:ss',
            'MM/DD/YYYY HH:mm:ss',
            'DD-MM-YYYY HH:mm',
            'MM-DD-YYYY HH:mm',
            'YYYY-MM-DD HH:mm'
        ];

        // Try parsing with each format
        for (const format of dateFormats) {
            try {
                const parsed = dayjs(cleanDate, format, true); // strict mode
                if (parsed.isValid()) {
                    const result = parsed.format('YYYY-MM-DD');
                    console.log(`✅ Successfully parsed "${cleanDate}" using format "${format}" → ${result}`);
                    return result;
                }
            } catch (error) {
                // Continue to next format
                continue;
            }
        }

        // Try flexible parsing without strict format (Day.js built-in parser)
        try {
            const flexibleParsed = dayjs(cleanDate);
            if (flexibleParsed.isValid()) {
                const result = flexibleParsed.format('YYYY-MM-DD');
                console.log(`✅ Successfully parsed "${cleanDate}" using flexible parsing → ${result}`);
                return result;
            }
        } catch (error) {
            console.error('❌ Flexible parsing failed:', error);
        }

        // Try with different locale assumptions
        const tryRegexPatterns = [
            // DD/MM/YYYY vs MM/DD/YYYY disambiguation
            {
                pattern: /^(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{4})$/,
                handler: (match) => {
                    const [, first, second, year] = match;
                    const firstNum = parseInt(first);
                    const secondNum = parseInt(second);
                    
                    // If first number > 12, it must be day (DD/MM/YYYY)
                    if (firstNum > 12) {
                        return dayjs(`${year}-${second.padStart(2, '0')}-${first.padStart(2, '0')}`);
                    }
                    // If second number > 12, it must be day (MM/DD/YYYY)
                    else if (secondNum > 12) {
                        return dayjs(`${year}-${first.padStart(2, '0')}-${second.padStart(2, '0')}`);
                    }
                    // Ambiguous - default to DD/MM/YYYY for international invoices
                    else {
                        return dayjs(`${year}-${second.padStart(2, '0')}-${first.padStart(2, '0')}`);
                    }
                }
            }
        ];

        for (const { pattern, handler } of tryRegexPatterns) {
            const match = cleanDate.match(pattern);
            if (match) {
                try {
                    const result = handler(match);
                    if (result && result.isValid && result.isValid()) {
                        const finalResult = result.format('YYYY-MM-DD');
                        console.log(`✅ Successfully parsed "${cleanDate}" using regex pattern → ${finalResult}`);
                        return finalResult;
                    }
                } catch (error) {
                    continue;
                }
            }
        }

        // Last resort: try to extract any date-like patterns
        const dateRegex = /\b(\d{1,2})\s*[-\/\.]\s*(\d{1,2})\s*[-\/\.]\s*(\d{2,4})\b/;
        const match = cleanDate.match(dateRegex);
        if (match) {
            const [, first, second, year] = match;
            const fullYear = year.length === 2 ? `20${year}` : year;
            
            try {
                // Try DD-MM-YYYY first (more common internationally)
                const attempt1 = dayjs(`${fullYear}-${second.padStart(2, '0')}-${first.padStart(2, '0')}`);
                if (attempt1.isValid()) {
                    const result = attempt1.format('YYYY-MM-DD');
                    console.log(`✅ Successfully parsed "${cleanDate}" using last resort DD-MM-YYYY → ${result}`);
                    return result;
                }
                
                // Try MM-DD-YYYY as fallback
                const attempt2 = dayjs(`${fullYear}-${first.padStart(2, '0')}-${second.padStart(2, '0')}`);
                if (attempt2.isValid()) {
                    const result = attempt2.format('YYYY-MM-DD');
                    console.log(`✅ Successfully parsed "${cleanDate}" using last resort MM-DD-YYYY → ${result}`);
                    return result;
                }
            } catch (error) {
                console.error('❌ Last resort parsing failed:', error);
            }
        }

        // Final fallback - log the failure and return empty string
        console.error('❌ Could not parse date in any format:', dateStr, '→ cleaned:', cleanDate);
        console.error('💡 Consider adding this format to the dateFormats array if it\'s a valid date');
        
        return ''; // Return empty string instead of today's date
    }

    // Helper method for validation
    static isValidDate(dateStr) {
        const parsed = this.parseInvoiceDate(dateStr);
        return parsed !== '';
    }
}

// 🆕 NEW JOURNAL VOUCHER PROCESSOR CLASS FOR PAYMENT TAB
class JournalVoucherProcessor {
    /**
     * Validates a journal voucher entry
     * @param {object} journalData - Journal voucher data
     * @returns {object} - Validation result
     */
    static validateJournalEntry(journalData) {
        const errors = [];
        const warnings = [];
        
        // Validate date
        if (!journalData.voucher_date) {
            errors.push('Voucher date is required');
        } else {
            const parsedDate = UniversalDateParser.parseInvoiceDate(journalData.voucher_date);
            if (!parsedDate) {
                errors.push('Invalid voucher date format');
            }
        }
        
        // Validate debit entries
        if (!journalData.debit_entries || !Array.isArray(journalData.debit_entries) || journalData.debit_entries.length === 0) {
            errors.push('At least one debit entry (expense ledger) is required');
        } else {
            journalData.debit_entries.forEach((entry, index) => {
                if (!entry.ledger_name || entry.ledger_name.trim().length === 0) {
                    errors.push(`Debit entry ${index + 1}: Ledger name is required`);
                }
                if (!entry.amount || parseFloat(entry.amount) <= 0) {
                    errors.push(`Debit entry ${index + 1}: Valid amount is required`);
                }
            });
        }
        
        // Validate credit entries
        if (!journalData.credit_entries || !Array.isArray(journalData.credit_entries) || journalData.credit_entries.length === 0) {
            errors.push('At least one credit entry (payment ledger) is required');
        } else {
            journalData.credit_entries.forEach((entry, index) => {
                if (!entry.ledger_name || entry.ledger_name.trim().length === 0) {
                    errors.push(`Credit entry ${index + 1}: Ledger name is required`);
                }
                if (!entry.amount || parseFloat(entry.amount) <= 0) {
                    errors.push(`Credit entry ${index + 1}: Valid amount is required`);
                }
            });
        }
        
        // Calculate totals and validate balance
        if (errors.length === 0) {
            const totalDebits = journalData.debit_entries.reduce((sum, entry) => sum + parseFloat(entry.amount), 0);
            const totalCredits = journalData.credit_entries.reduce((sum, entry) => sum + parseFloat(entry.amount), 0);
            
            const difference = Math.abs(totalDebits - totalCredits);
            
            if (difference > 0.01) { // Allow for small rounding differences
                errors.push(`Journal entries must balance. Debit total: ₹${totalDebits.toFixed(2)}, Credit total: ₹${totalCredits.toFixed(2)}, Difference: ₹${difference.toFixed(2)}`);
            }
            
            if (totalDebits === 0 || totalCredits === 0) {
                errors.push('Both debit and credit amounts must be greater than zero');
           }
       }
       
       // Validate narration
       if (!journalData.narration || journalData.narration.trim().length < 3) {
           warnings.push('Consider adding a more descriptive narration for the journal entry');
       }
       
       return {
           isValid: errors.length === 0,
           errors: errors,
           warnings: warnings
       };
   }
   
   /**
    * Processes and formats journal voucher data for database storage
    * @param {object} journalData - Raw journal voucher data
    * @returns {object} - Processed journal voucher data
    */
   static processJournalEntry(journalData) {
      const processedData = {
          voucher_date: UniversalDateParser.parseInvoiceDate(journalData.voucher_date),
          voucher_type: 'Journal',
          narration: journalData.narration ? journalData.narration.trim() : '',
          entry_type: 'payment_journal',
          debit_entries: [],
          credit_entries: [],
          total_amount: 0,
          status: 'pending_review'
      };
      
      // Process debit entries
      journalData.debit_entries.forEach((entry, index) => {
          processedData.debit_entries.push({
              ledger_name: entry.ledger_name.trim(),
              amount: this.roundToTwoDecimals(parseFloat(entry.amount)),
              entry_type: 'debit',
              sequence: index + 1
          });
      });
      
      // Process credit entries
      journalData.credit_entries.forEach((entry, index) => {
          processedData.credit_entries.push({
              ledger_name: entry.ledger_name.trim(),
              amount: this.roundToTwoDecimals(parseFloat(entry.amount)),
              entry_type: 'credit',
              sequence: index + 1
          });
      });
      
      // Calculate total amount (sum of debits or credits - they should be equal)
      processedData.total_amount = this.roundToTwoDecimals(
          processedData.debit_entries.reduce((sum, entry) => sum + entry.amount, 0)
      );
      
      return processedData;
  }
  
  /**
   * Converts journal voucher to expense table format for unified storage
   * @param {object} journalData - Processed journal voucher data
   * @returns {object} - Data formatted for expenses table
   */
  static convertToExpenseFormat(journalData) {
      // For Journal vouchers, we'll use the first debit entry as the primary description
      const primaryDebitEntry = journalData.debit_entries[0];
      const primaryCreditEntry = journalData.credit_entries[0];
      
      // Create a readable invoice number for journal entries
      const journalNumber = this.generateJournalNumber(journalData.voucher_date);
      
      return {
          invoice_date: journalData.voucher_date,
          invoice_number: journalNumber,
          vendor_name: `Journal Entry - ${primaryCreditEntry.ledger_name}`,
          vendor_gstn: '',
          customer_name: '',
          customer_gstn: '',
          place_of_supply: '',
          taxable_amount: journalData.total_amount,
          igst_amount: 0,
          cgst_amount: 0,
          sgst_amount: 0,
          cess_amount: 0,
          round_off: 0,
          invoice_value: journalData.total_amount,
          tds_rate: 0,
          tds_amount: 0,
          description: `${primaryDebitEntry.ledger_name} - ${journalData.narration}`,
          hsn_sac_code: '',
          line_item_amount: journalData.total_amount,
          quantity: 1,
          unit_rate: journalData.total_amount,
          voucher_type: 'Journal',
          ledger_name: primaryDebitEntry.ledger_name,
          vendor_address: '',
          total_tax_amount: 0,
          cgst_rate: 0,
          sgst_rate: 0,
          igst_rate: 0,
          cess_rate: 0,
          category: 'Journal Entry',
          file_path: '',
          original_filename: '',
          file_hash: '',
          parseur_document_id: null,
          extracted_text: JSON.stringify({
              debit_entries: journalData.debit_entries,
              credit_entries: journalData.credit_entries,
              narration: journalData.narration
          }),
          confidence_score: 1.0, // Manual entries have 100% confidence
          amount_confidence: 1.0,
          document_type: 'journal',
          processing_time_ms: 0,
          status: 'pending_review',
          processing_source: 'manual_journal',
          line_items: JSON.stringify(this.createLineItemsFromJournal(journalData)),
          line_items_count: journalData.debit_entries.length + journalData.credit_entries.length,
          has_line_items: true,
          table_structure_confidence: 1.0,
          processing_method: 'manual',
          validation_errors: '[]',
          validation_warnings: '[]',
          entry_type: 'payment_journal'
      };
  }
  
  /**
   * Creates line items representation of journal entries
   * @param {object} journalData - Processed journal voucher data
   * @returns {array} - Line items array
   */
  static createLineItemsFromJournal(journalData) {
      const lineItems = [];
      
      // Add debit entries as line items
      journalData.debit_entries.forEach((entry, index) => {
          lineItems.push({
              description: `${entry.ledger_name} (Debit)`,
              hsn_code: '',
              quantity: 1,
              unit_rate: entry.amount,
              line_total: entry.amount,
              tax_rate: 0,
              tax_amount: 0,
              entry_type: 'debit',
              sequence: index + 1
          });
      });
      
      // Add credit entries as line items
      journalData.credit_entries.forEach((entry, index) => {
          lineItems.push({
              description: `${entry.ledger_name} (Credit)`,
              hsn_code: '',
              quantity: 1,
              unit_rate: -entry.amount, // Negative for credit
              line_total: -entry.amount,
              tax_rate: 0,
              tax_amount: 0,
              entry_type: 'credit',
              sequence: index + 1
          });
      });
      
      return lineItems;
  }
  
  /**
   * Generates a unique journal number
   * @param {string} date - Voucher date
   * @returns {string} - Generated journal number
   */
  static generateJournalNumber(date) {
      const dateObj = dayjs(date);
      const year = dateObj.format('YYYY');
      const month = dateObj.format('MM');
      const day = dateObj.format('DD');
      const timestamp = Date.now().toString().slice(-4);
      
      return `JV${year}${month}${day}-${timestamp}`;
 }
 
 /**
  * Rounds number to two decimal places
  * @param {number} value - Number to round
  * @returns {number} - Rounded number
  */
 static roundToTwoDecimals(value) {
     return Math.round((value + Number.EPSILON) * 100) / 100;
 }
}

// 🆕 INTEGRATED TALLY IMPORT FUNCTIONALITY - FIXED TIMING FOR LEDGER STANDARDIZATION
class TallyIntegrator {
 /**
  * Groups data by voucher number
  * @param {Array} data - Excel data array
  * @returns {Object} - Grouped data by voucher number
  */
 static groupByVoucher(data) {
     const grouped = {};
     data.forEach((entry) => {
         const key = `${entry["Voucher Number"]}`;
         if (!grouped[key]) {
             grouped[key] = [];
         }
         grouped[key].push(entry);
     });
     return grouped;
 }

 /**
  * Extract unique ledger names from data
  * @param {Array} data - Excel data array
  * @returns {Array} - Array of unique ledger names
  */
 static extractLedgerNames(data) {
     const ledgers = new Set();
     data.forEach(entry => {
         if (entry["Ledger Name"] && entry["Ledger Name"].trim()) {
             ledgers.add(entry["Ledger Name"].trim());
         }
     });
     return Array.from(ledgers);
 }

 /**
  * Extract unique stock items from data
  * @param {Array} data - Excel data array
  * @returns {Object} - Object with items array and details
  */
 static extractStockItems(data) {
     const stockItems = new Set();
     const stockDetails = {};
     
     data.forEach(entry => {
         const itemName = entry["Item Name"];
         const quantity = parseFloat(entry["Billed Quantity"]) || 0;
         const rate = parseFloat(entry["Item Rate"]) || 0;
         const unit = entry["Item Rate per"] || 'Nos';
         
         if (itemName && itemName.trim() && quantity > 0) {
             const cleanItemName = itemName.trim();
             stockItems.add(cleanItemName);
             
             // Store additional details for stock item creation
             if (!stockDetails[cleanItemName]) {
                 stockDetails[cleanItemName] = {
                     name: cleanItemName,
                     unit: unit,
                     rate: rate,
                     hasStock: true
                 };
             }
         }
     });
     
     return {
         items: Array.from(stockItems),
         details: stockDetails
     };
 }

 /**
  * Build stock items XML for Tally import
  * @param {Array} stockItems - Array of stock item names
  * @param {Object} stockDetails - Stock item details
  * @returns {String} - XML string for stock items
  */
 static buildStockItemsXML(stockItems, stockDetails) {
     const items = stockItems.map(itemName => {
         const details = stockDetails[itemName] || {};
         
         return {
             STOCKITEM: {
                 ACTION: 'Create',
                 NAME: itemName,
                 PARENT: 'Primary',
                 TAXCLASSIFICATIONNAME: '',
                 TAXTYPE: 'Others',
                 ISACTIVE: 'Yes',
                 ISCOSTCENTREON: 'No',
                 ISBATCHWISEON: 'No',
                 ISPERISHABLEON: 'No',
                 ISMAINTAINBALANCE: 'Yes',
                 ISSTOCKITEM: 'Yes',
                 HASMFGDATE: 'No',
                 ALLOWUSEOFEXPIREDITEMS: 'No',
                 IGNOREPHYSICALDIFFERENCE: 'No',
                 IGNORENEGATIVESTOCK: 'No',
                 TREATSALESASMANUFACTURED: 'No',
                 TREATPURCHASEASCONSUMED: 'No',
                 TREATREJECTSASSCRAP: 'No',
                 HASBATCHES: 'No',
                 HASWARRANTED: 'No',
                 TRACKED: 'No',
                 PERISHABLE: 'No',
                 ENTRYTYPE: 'Item',
                 COSTINGMETHOD: 'Avg. Cost',
                 VALUATIONMETHOD: 'Avg. Price',
                 BASEUNITS: details.unit || 'Nos',
                 ADDITIONALUNITS: '',
                 OPENINGBALANCE: '0',
                 OPENINGVALUE: '0',
                 OPENINGRATE: details.rate || '0'
             }
         };
     });

     const root = {
         ENVELOPE: {
             HEADER: {
                 TALLYREQUEST: 'Import Data'
             },
             BODY: {
                 IMPORTDATA: {
                     REQUESTDESC: {
                         REPORTNAME: 'All Masters',
                         STATICVARIABLES: {
                             SVCURRENTCOMPANY: process.env.TALLY_COMPANY_NAME || 'Topaz International (2021-22)'
                         }
                     },
                     REQUESTDATA: {
                         'TALLYMESSAGE': items
                     }
                 }
             }
         }
     };

     const builder = new XMLBuilder({ 
         ignoreAttributes: false,
         format: true,
         suppressEmptyNode: true
     });
     
     return builder.build(root);
 }

 /**
  * Build ledgers XML for Tally import with smart categorization
  * @param {Array} ledgerNames - Array of ledger names
  * @param {Object} ledgerConfig - Optional ledger configuration
  * @returns {String} - XML string for ledgers
  */
 static buildAdvancedLedgersXML(ledgerNames, ledgerConfig = {}) {
     const ledgers = ledgerNames.map(name => {
         let parent = 'Sundry Debtors';
         
         const lowerName = name.toLowerCase();
         if (lowerName.includes('bank') || lowerName.includes('cash')) {
             parent = 'Bank Accounts';
         } else if (lowerName.includes('expense') || lowerName.includes('cost')) {
             parent = 'Direct Expenses';
         } else if (lowerName.includes('purchase')) {
             parent = 'Purchase Accounts';
         } else if (lowerName.includes('income') || lowerName.includes('revenue')) {
             parent = 'Direct Incomes';
         } else if (lowerName.includes('sales')) {
             parent = 'Sales Accounts';
         } else if (lowerName.includes('liability') || lowerName.includes('payable')) {
             parent = 'Sundry Creditors';
         } else if (lowerName.includes('asset') || lowerName.includes('receivable')) {
             parent = 'Sundry Debtors';
         } else if (lowerName.includes('igst') || lowerName.includes('cgst') || lowerName.includes('sgst') || lowerName.includes('gst')) {
             parent = 'Duties & Taxes';
         }
         
         if (ledgerConfig[name]) {
             parent = ledgerConfig[name].parent || parent;
         }

         return {
             LEDGER: {
                 ACTION: 'Create',
                 NAME: name,
                 PARENT: parent,
                 ISBILLWISEON: 'No',
                 ISCOSTCENTREON: 'No',
                 ISINTERESTON: 'No',
                 ALLOWINMOBILE: 'No',
                 ISCOSTTRACKINGON: 'No',
                 ISBENEFICIARYCODEON: 'No',
                 ISUPDATINGTARGETID: 'No',
                 ASORIGINAL: 'Yes',
                 AFFECTSSTOCK: 'No',
                 USEFORVAT: 'No',
                 AUDITED: 'No',
                 FORPAYROLL: 'No',
                 ISTDSDEDUCTEE: 'No',
                 ISTCSAPPLICABLE: 'No',
                 ISTDSTAXABLE: 'No',
                 ISGSTAPPLICABLE: 'No',
                 ISGSTTAXABLE: 'No',
                 ISINPUTCREDIT: 'No',
                 ISEXEMPTED: 'No',
                 ISSYSTEM: 'No',
                 ISFIXED: 'No',
                 ISCONDENSED: 'No'
             }
         };
     });

     const root = {
         ENVELOPE: {
             HEADER: {
                 TALLYREQUEST: 'Import Data'
             },
             BODY: {
                 IMPORTDATA: {
                     REQUESTDESC: {
                         REPORTNAME: 'All Masters',
                         STATICVARIABLES: {
                             SVCURRENTCOMPANY: process.env.TALLY_COMPANY_NAME || 'Topaz International (2021-22)'
                         }
                     },
                     REQUESTDATA: {
                         'TALLYMESSAGE': ledgers
                     }
                 }
             }
         }
     };

     const builder = new XMLBuilder({ 
         ignoreAttributes: false,
         format: true,
         suppressEmptyNode: true
     });
     
     return builder.build(root);
 }

 /**
  * Format date for Tally (YYYYMMDD format)
  * @param {String} dateStr - Date string to format
  * @returns {String} - Formatted date string
  */
 static formatTallyDate(dateStr) {
     if (!dateStr) {
         const today = new Date();
         const day = String(today.getDate()).padStart(2, '0');
         const month = String(today.getMonth() + 1).padStart(2, '0');
         const year = today.getFullYear();
         return `${year}${month}${day}`;
     }
     
     let date;
     
     if (typeof dateStr === 'string') {
         if (dateStr.includes('/')) {
             const parts = dateStr.split('/');
             if (parts.length === 3) {
                 const day = parts[0].padStart(2, '0');
                 const month = parts[1].padStart(2, '0');
                 const year = parts[2];
                 return `${year}${month}${day}`;
             }
         }
         date = new Date(dateStr);
     } else if (typeof dateStr === 'number') {
         date = new Date((dateStr - 25569) * 86400 * 1000);
     } else if (dateStr instanceof Date) {
         date = dateStr;
     } else {
         return '';
     }
     
     if (isNaN(date.getTime())) {
         console.warn(`Invalid date format: ${dateStr}`);
         const today = new Date();
         const day = String(today.getDate()).padStart(2, '0');
         const month = String(today.getMonth() + 1).padStart(2, '0');
         const year = today.getFullYear();
         return `${year}${month}${day}`;
     }
     
     const day = String(date.getDate()).padStart(2, '0');
     const month = String(date.getMonth() + 1).padStart(2, '0');
     const year = date.getFullYear();
     
     return `${year}${month}${day}`;
 }

 /**
  * Build vouchers XML for Tally import with inventory allocations
  * @param {Object} groupedData - Grouped voucher data
  * @param {Object} stockDetails - Stock item details
  * @returns {String} - XML string for vouchers
   */
static buildEnhancedTallyXML(groupedData, stockDetails) {
  const vouchers = [];

  Object.keys(groupedData).forEach((voucherNumber) => {
      const entries = groupedData[voucherNumber];
      const first = entries[0];

      let totalDr = 0, totalCr = 0;
      const allEntries = [];

      console.log(`\n📝 Processing Voucher: ${voucherNumber}`);

      entries.forEach((e, index) => {
// 🔍 DEBUG: Check what data we're getting
      console.log(`🔍 DEBUG Row ${index}:`, {
          ledgerName: e["Ledger Name"],
          itemName: e["Item Name"], 
          quantity: e["Billed Quantity"],
          rate: e["Item Rate"],
          amount: e["Item Amount"],
          unit: e["Item Rate per"]
      });
          const amount = parseFloat(e["Ledger Amount"]) || 0;
          const drcr = (e["Ledger Amount Dr/Cr"] || '').toLowerCase().trim();

          console.log(`   Entry ${index + 1}: ${e["Ledger Name"]} - ${amount} ${drcr.toUpperCase()}`);

          if (amount === 0 || !e["Ledger Name"] || e["Ledger Name"].trim() === '') {
              console.warn(`   ⚠️ Skipping invalid entry: ${e["Ledger Name"]}`);
              return;
          }

          if (drcr !== 'dr' && drcr !== 'cr') {
              console.warn(`   ⚠️ Invalid Dr/Cr value: '${drcr}' for ledger: ${e["Ledger Name"]}`);
              return;
          }

          const ledgerEntry = {
              LEDGERNAME: e["Ledger Name"].trim(),
              ISDEEMEDPOSITIVE: drcr === 'dr' ? 'Yes' : 'No',
              AMOUNT: drcr === 'dr' ? (-amount).toFixed(2) : amount.toFixed(2),
          };

console.log(`🧾 DEBUG: Adding Ledger Entry:
  Voucher: ${voucherNumber}
  → Original Name: "${e["Ledger Name"]}"
  → Standardized Name: "${ledgerEntry.LEDGERNAME}"
  → Amount: ${ledgerEntry.AMOUNT}
  → Dr/Cr: ${drcr.toUpperCase()}
`);

          // 🔧 FIX: Check if this is a stock-affecting ledger and has stock items
          const itemName = e["Item Name"];
          const quantity = parseFloat(e["Billed Quantity"]) || 0;
          const rate = parseFloat(e["Item Rate"]) || 0;
          const unit = e["Item Rate per"] || 'Nos';
          const itemAmount = parseFloat(e["Item Amount"]) || 0;

          // 🎯 KEY FIX: Only add inventory allocations for Purchase/Sales ledgers with stock items
          if (itemName && itemName.trim() && quantity > 0 && rate > 0) {
              console.log(`   📦 Adding Stock Item: ${itemName} - Qty: ${quantity} ${unit} @ Rate: ${rate}`);
              
              // 🔧 FIXED: Proper inventory allocation structure
              const inventoryAllocation = {
                  STOCKITEMNAME: itemName.trim(),
                  ISDEEMEDPOSITIVE: drcr === 'dr' ? 'Yes' : 'No',
                  ISLASTDEEMEDPOSITIVE: drcr === 'dr' ? 'Yes' : 'No',
                  ISAUTONEGATE: 'No',
                  ISCAPVATTAXALTERED: 'No',
                  ISCAPVATNOTCLAIMED: 'No',
                  AMOUNT: drcr === 'dr' ? (-itemAmount).toFixed(2) : itemAmount.toFixed(2),
                  ACTUALQTY: drcr === 'dr' ? `${quantity} ${unit}` : `${-quantity} ${unit}`,
                  BILLEDQTY: drcr === 'dr' ? `${quantity} ${unit}` : `${-quantity} ${unit}`,
                  RATE: `${rate}/${unit}`,
                  DISCOUNT: '0'
              };

              // 🎯 KEY FIX: Add inventory allocation to the ledger entry
              ledgerEntry['INVENTORYALLOCATIONS.LIST'] = [inventoryAllocation];
              
              console.log(`   ✅ Inventory allocation added for ${itemName}`);
          }

          console.log(`   → XML: ${ledgerEntry.LEDGERNAME} = ${ledgerEntry.AMOUNT} (ISDEEMEDPOSITIVE: ${ledgerEntry.ISDEEMEDPOSITIVE})`);

          drcr === 'dr' ? totalDr += amount : totalCr += amount;
          allEntries.push(ledgerEntry);
      });

      console.log(`   📊 Totals: Dr = ${totalDr}, Cr = ${totalCr}`);

      // 🎯 NEW FIX: Reorder entries for Purchase vouchers (Credit entries first)
      const voucherTypeName = first["Voucher Type Name"] || 'Purchase';
      
      if (voucherTypeName === 'Purchase') {
          console.log(`   🔄 Reordering entries for Purchase voucher: Credit entries first`);
          
          // Separate Credit and Debit entries
          const creditEntries = allEntries.filter(entry => entry.ISDEEMEDPOSITIVE === 'No');
          const debitEntries = allEntries.filter(entry => entry.ISDEEMEDPOSITIVE === 'Yes');
          
          // Reorder: Credit entries first, then Debit entries
          allEntries.length = 0; // Clear the array
          allEntries.push(...creditEntries, ...debitEntries);
          
          console.log(`   ✅ Purchase voucher reordered: ${creditEntries.length} Credit entries + ${debitEntries.length} Debit entries`);
          console.log(`   📋 Entry order: ${allEntries.map(e => `${e.LEDGERNAME} (${e.ISDEEMEDPOSITIVE === 'Yes' ? 'Dr' : 'Cr'})`).join(' → ')}`);
      }

      // Handle Round Off if there's a small difference
      const diff = parseFloat((totalDr - totalCr).toFixed(2));
      if (Math.abs(diff) >= 0.01) {
          console.log(`   🔄 Adding Round Off: ${diff}`);
          
          const roundOffEntry = {
              LEDGERNAME: 'Round Off',
              ISDEEMEDPOSITIVE: diff > 0 ? 'No' : 'Yes',
              AMOUNT: diff > 0 ? diff.toFixed(2) : (-diff).toFixed(2)
          };
          
          // 🎯 NEW FIX: For Purchase vouchers, add Round Off based on entry type
          if (voucherTypeName === 'Purchase') {
              // For Purchase vouchers, add Round Off with other entries of the same type
              if (roundOffEntry.ISDEEMEDPOSITIVE === 'No') {
                  // Credit entry - add after other credit entries but before debit entries
                  const creditCount = allEntries.filter(e => e.ISDEEMEDPOSITIVE === 'No').length;
                  allEntries.splice(creditCount, 0, roundOffEntry);
              } else {
                  // Debit entry - add at the end
                  allEntries.push(roundOffEntry);
              }
          } else {
              // For other voucher types, add at the end (existing behavior)
              allEntries.push(roundOffEntry);
          }
      }

      if (allEntries.length === 0) {
          console.warn(`   ⚠️ No valid entries for voucher ${voucherNumber}`);
          return;
      }

      const voucherObj = {
          REMOTEID: '',
          VCHKEY: '',
          VCHTYPE: voucherTypeName,
          ACTION: 'Create',
          OBJVIEW: 'Accounting Voucher View',
          DATE: this.formatTallyDate(first["Voucher Date"]),
          NARRATION: first["Voucher Narration"] || '',
          VOUCHERTYPENAME: voucherTypeName,
          VOUCHERNUMBER: voucherNumber,
          REFERENCE: first["Voucher Number"] || '',
          'ALLLEDGERENTRIES.LIST': allEntries
      };

      vouchers.push(voucherObj);
      console.log(`   ✅ Voucher ${voucherNumber} (${voucherTypeName}) prepared with ${allEntries.length} entries in correct order`);

  }); // 🔧 FIX: Added missing closing brace and parenthesis for Object.keys(groupedData).forEach()

  const root = {
      ENVELOPE: {
          HEADER: {
              TALLYREQUEST: 'Import Data'
          },
          BODY: {
              IMPORTDATA: {
                  REQUESTDESC: {
                      REPORTNAME: 'Vouchers',
                      STATICVARIABLES: {
                          SVCURRENTCOMPANY: process.env.TALLY_COMPANY_NAME || 'Topaz International (2021-22)'
                      }
                  },
                  REQUESTDATA: {
                      'TALLYMESSAGE': vouchers.map(v => ({ VOUCHER: v }))
                  }
              }
          }
      }
  };

  const builder = new XMLBuilder({ 
      ignoreAttributes: false,
      format: true,
      suppressEmptyNode: true
  });
  
  return builder.build(root);
}

 /**
  * Send XML to Tally with detailed response handling
  * @param {String} xmlData - XML data to send
  * @param {String} description - Description for logging
  * @returns {Promise} - Response from Tally
  */
 static async sendToTally(xmlData, description) {
     try {
         const tallyUrl = process.env.TALLY_URL || 'http://localhost:9000';
         
         const response = await axios.post(tallyUrl, xmlData, {
             headers: { 'Content-Type': 'application/xml' },
             timeout: 30000
         });
         
         console.log(`✅ ${description} sent to Tally:`, response.status);
         console.log('Response:', response.data);
         
         if (response.data.includes('CREATED>') && response.data.includes('EXCEPTIONS>')) {
             const createdMatch = response.data.match(/<CREATED>(\d+)<\/CREATED>/);
             const exceptionsMatch = response.data.match(/<EXCEPTIONS>(\d+)<\/EXCEPTIONS>/);
             
             if (createdMatch && exceptionsMatch) {
                 const created = parseInt(createdMatch[1]);
                 const exceptions = parseInt(exceptionsMatch[1]);
                 
                 console.log(`📊 Results: ${created} created, ${exceptions} exceptions`);
                 
                 if (exceptions > 0) {
                     console.log('⚠️ Some entries had exceptions. Check Tally for details.');
                 }
                 
                 return { created, exceptions, response: response.data };
             }
         }
         
         return { created: 0, exceptions: 0, response: response.data };
     } catch (err) {
         console.error(`❌ Error sending ${description}:`, err.message);
         if (err.response) {
             console.error('Response data:', err.response.data);
         }
         throw err;
     }
 }

 /**
  * 🎯 MAIN IMPORT FUNCTION WITH FIXED LEDGER STANDARDIZATION TIMING
  * @param {String} excelFilePath - Path to the Excel file
  * @param {LedgerStandardizer} ledgerStandardizer - Optional ledger standardizer instance
  * @returns {Promise<Object>} - Import results
  */
 static async importToTally(excelFilePath, ledgerStandardizer = null) {
     try {
         console.log('🚀 Starting Tally import process...');
         
         if (!fs.existsSync(excelFilePath)) {
             throw new Error(`Excel file not found: ${excelFilePath}`);
         }
         
         const workbook = XLSX.readFile(excelFilePath);
         const sheetName = workbook.SheetNames[0];
         const sheet = workbook.Sheets[sheetName];
         let data = XLSX.utils.sheet_to_json(sheet);

         console.log(`📊 Loaded ${data.length} rows from Excel`);

         if (data.length === 0) {
             throw new Error('No data found in Excel file');
         }

         // 🎯 CRITICAL FIX: APPLY LEDGER STANDARDIZATION HERE AFTER LOADING BUT BEFORE PROCESSING
         if (ledgerStandardizer && ledgerStandardizer.initialized) {
             console.log('🔄 Applying ledger name standardization to loaded Excel data...');
             data = ledgerStandardizer.standardizeTallyExportData(data);
             console.log('✅ Ledger standardization applied to Excel data before processing');
         } else {
             console.log('⚠️ Ledger standardization not available - processing with original names');
         }

         const results = {
             totalRows: data.length,
             stockItems: { created: 0, exceptions: 0 },
             ledgers: { created: 0, exceptions: 0 },
             vouchers: { created: 0, exceptions: 0 },
             errors: [],
             ledgerStandardization: {
                 applied: !!(ledgerStandardizer && ledgerStandardizer.initialized),
                 statistics: ledgerStandardizer ? ledgerStandardizer.getStatistics() : null
             }
         };

         // Step 1: Extract and handle stock items
         const stockInfo = this.extractStockItems(data);
         console.log(`📦 Found ${stockInfo.items.length} unique stock items`);

         if (stockInfo.items.length > 0) {
             try {
                 const stockItemsXML = this.buildStockItemsXML(stockInfo.items, stockInfo.details);
                 const stockResult = await this.sendToTally(stockItemsXML, 'Stock Items');
                 results.stockItems = stockResult;
                 
                 console.log('⏳ Waiting for stock items to be processed...');
                 await new Promise(resolve => setTimeout(resolve, 3000));
             } catch (error) {
                 console.error('❌ Stock items import failed:', error.message);
                 results.errors.push(`Stock items: ${error.message}`);
             }
         }

         // Step 2: Extract and create ledgers (now with standardized names)
         const ledgerNames = this.extractLedgerNames(data);
         console.log(`🏦 Found ${ledgerNames.length} unique ledgers (after standardization)`);

         if (ledgerNames.length > 0) {
             try {
                 const ledgerConfig = {
                     'Cash': { parent: 'Cash-in-Hand' },
                     'Bank Account': { parent: 'Bank Accounts' },
                     'Purchase A/c': { parent: 'Purchase Accounts' },
                     'Sales': { parent: 'Sales Accounts' },
                     'Purchase': { parent: 'Purchase Accounts' },
                     'Round Off': { parent: 'Direct Expenses' },
                     'IGST 18%': { parent: 'Duties & Taxes' },
                     'CGST 9%': { parent: 'Duties & Taxes' },
                     'SGST 9%': { parent: 'Duties & Taxes' },
                     'GST 18%': { parent: 'Duties & Taxes' },
                     // 🎯 Add standardized ledger configurations
                     'CGST INPUT': { parent: 'Duties & Taxes' },
                     'SGST INPUT': { parent: 'Duties & Taxes' },
                     'IGST INPUT': { parent: 'Duties & Taxes' },
                     'CGST OUTPUT': { parent: 'Duties & Taxes' },
                     'SGST OUTPUT': { parent: 'Duties & Taxes' },
                     'IGST OUTPUT': { parent: 'Duties & Taxes' }
                 };

                 const ledgersXML = this.buildAdvancedLedgersXML(ledgerNames, ledgerConfig);
                 const ledgerResult = await this.sendToTally(ledgersXML, 'Ledgers');
                 results.ledgers = ledgerResult;
                 
                 console.log('⏳ Waiting for ledgers to be processed...');
                 await new Promise(resolve => setTimeout(resolve, 2000));
             } catch (error) {
                 console.error('❌ Ledgers import failed:', error.message);
                 results.errors.push(`Ledgers: ${error.message}`);
             }
         }

         // Step 3: Create vouchers with inventory allocations (using standardized data)
         const grouped = this.groupByVoucher(data);
         console.log(`📋 Found ${Object.keys(grouped).length} unique vouchers`);

         if (Object.keys(grouped).length > 0) {
             try {
                 const vouchersXML = this.buildEnhancedTallyXML(grouped, stockInfo.details);
                 const voucherResult = await this.sendToTally(vouchersXML, 'Vouchers with Stock Items');
                 results.vouchers = voucherResult;
             } catch (error) {
                 console.error('❌ Vouchers import failed:', error.message);
                 results.errors.push(`Vouchers: ${error.message}`);
             }
         }

         console.log('🎉 Tally import process completed!');
         return results;

     } catch (error) {
         console.error('❌ Tally import process failed:', error.message);
         throw error;
     }
 }
}

// Configuration (kept private/internal)
const CLAUDE_CONFIG = {
 apiKey: process.env.ANTHROPIC_API_KEY,
 model: 'claude-sonnet-4-6',
 enabled: !!(process.env.ANTHROPIC_API_KEY)
};

const PARSEUR_CONFIG = {
 apiToken: process.env.PARSEUR_API_TOKEN,
 baseUrl: 'https://api.parseur.com',
 mailboxId: process.env.PARSEUR_MAILBOX_ID,
 webhookSecret: process.env.PARSEUR_WEBHOOK_SECRET,
 enabled: !!(process.env.PARSEUR_API_TOKEN && process.env.PARSEUR_MAILBOX_ID)
};

// Initialize services
const documentAIClient = new DocumentProcessorServiceClient({
 keyFilename: process.env.GOOGLE_APPLICATION_CREDENTIALS || './google-credentials.json'
});

const DOCUMENT_AI_CONFIG = {
 projectId: process.env.GOOGLE_CLOUD_PROJECT_ID,
 location: process.env.DOCUMENT_AI_LOCATION || 'us',
 processorId: process.env.DOCUMENT_AI_PROCESSOR_ID,
 processorVersion: process.env.DOCUMENT_AI_PROCESSOR_VERSION || 'rc'
};

const openai = new OpenAI({
 apiKey: process.env.OPENAI_API_KEY
});

const anthropic = new Anthropic({
 apiKey: process.env.ANTHROPIC_API_KEY
});

if (!fs.existsSync('./uploads')) fs.mkdirSync('./uploads');
if (!fs.existsSync('./processed')) fs.mkdirSync('./processed');
if (!fs.existsSync('./tally-exports')) fs.mkdirSync('./tally-exports'); // 🆕 Directory for Tally Excel files

const ALLOWED_EXTENSIONS = /\.(jpeg|jpg|png|gif|bmp|tiff|webp|pdf)$/i;

const upload = multer({
 storage: multer.diskStorage({
     destination: (req, file, cb) => cb(null, './uploads'),
     filename: (req, file, cb) => cb(null, Date.now() + '-' + Math.round(Math.random() * 1E9) + path.extname(file.originalname))
 }),
 limits: { fileSize: 25 * 1024 * 1024 },
 // NOTE: Do NOT use cb(null, false) here — in multer v1.4.5-lts.1 that throws
 // LIMIT_UNEXPECTED_FILE instead of silently skipping. Validate extensions in the route handler.
 fileFilter: (req, file, cb) => cb(null, true)
});

let db;

// 🆕 Initialize Ledger Standardizer
const ledgerStandardizer = new LedgerStandardizer();

// FIXED: Helper functions for safe value preparation
const prepareValue = (value, fieldName) => {
 if (value === undefined || value === null) {
     console.log(`⚠️ Field ${fieldName} is ${value}, using default empty string`);
     return '';
 }
 return value.toString();
};

const prepareDateValue = (value, fieldName) => {
 if (value === undefined || value === null || value === '') {
     console.log(`⚠️ Date field ${fieldName} is empty, using empty string`);
     return ''; // Don't default to current date
 }
 
 // Use the universal parser
 const parsedDate = UniversalDateParser.parseInvoiceDate(value);
 if (parsedDate === '') {
     console.log(`⚠️ Could not parse date field ${fieldName}: ${value}, using empty string`);
     return '';
 }
 
 return parsedDate;
};

const prepareNumericValue = (value, fieldName, defaultValue = 0) => {
 if (value === undefined || value === null || isNaN(value)) {
     console.log(`⚠️ Numeric field ${fieldName} is ${value}, using ${defaultValue}`);
     return defaultValue;
 }
 const numValue = parseFloat(value);
 return isNaN(numValue) ? defaultValue : numValue;
};

const prepareBooleanValue = (value, fieldName) => {
 if (value === undefined || value === null) {
     console.log(`⚠️ Boolean field ${fieldName} is ${value}, using 0`);
     return 0;
 }
 return value ? 1 : 0;
};

// 🔧 ENHANCED LINE ITEMS VALIDATION AND PROCESSING
class LineItemsValidator {
 /**
  * Validates a single line item
  * @param {object} item - Line item object
  * @param {number} index - Index for error messages
  * @returns {object} - Validation result
  */
 static validateLineItem(item, index) {
     const errors = [];
     
     // Required fields validation
     if (!item.description || item.description.trim().length < 3) {
         errors.push(`Line item ${index + 1}: Description is required (minimum 3 characters)`);
     }
     
     if (!item.quantity || item.quantity <= 0) {
         errors.push(`Line item ${index + 1}: Quantity must be greater than 0`);
     }
     
     if (item.unit_rate === undefined || item.unit_rate === null || item.unit_rate === '' || parseFloat(item.unit_rate) < 0) {
         errors.push(`Line item ${index + 1}: Unit rate must be 0 or greater`);
     }
     
     if (item.tax_rate === undefined || item.tax_rate < 0 || item.tax_rate > 100) {
         errors.push(`Line item ${index + 1}: Tax rate must be between 0 and 100`);
     }
     
     // Optional HSN code validation
     if (item.hsn_code && !/^\d{4,8}$/.test(item.hsn_code)) {
         errors.push(`Line item ${index + 1}: HSN code must be 4-8 digits`);
     }
     
     return {
         isValid: errors.length === 0,
         errors: errors
     };
 }
 
 /**
  * Processes and calculates line item values
  * @param {object} item - Raw line item
  * @returns {object} - Processed line item
  */
 static processLineItem(item) {
     const quantity = parseFloat(item.quantity) || 1;
     const unitRate = parseFloat(item.unit_rate) ?? 0;
     const taxRate = parseFloat(item.tax_rate) ?? 0;

     // When rate is 0, preserve the manually entered line_total (e.g. energy charges billed by formula)
     const calculatedTotal = this.roundToTwoDecimals(quantity * unitRate);
     const lineTotal = (unitRate === 0 && item.line_total > 0)
         ? this.roundToTwoDecimals(parseFloat(item.line_total))
         : calculatedTotal;

     // Calculate tax amount
     const taxAmount = this.roundToTwoDecimals((lineTotal * taxRate) / 100);
     
     return {
         description: item.description.trim(),
         hsn_code: item.hsn_code || '',
         quantity: quantity,
         unit_rate: unitRate,
         line_total: lineTotal,
         tax_rate: taxRate,
         tax_amount: taxAmount
     };
 }
 
 /**
  * Validates and processes array of line items
  * @param {array} lineItems - Array of line items
  * @returns {object} - Validation and processing result
  */
 static validateAndProcessLineItems(lineItems) {
     if (!Array.isArray(lineItems)) {
         return {
             isValid: false,
             errors: ['Line items must be an array'],
             processedItems: []
         };
     }
     
     const allErrors = [];
     const processedItems = [];
     
     lineItems.forEach((item, index) => {
         const validation = this.validateLineItem(item, index);
         
         if (!validation.isValid) {
             allErrors.push(...validation.errors);
         } else {
             processedItems.push(this.processLineItem(item));
         }
     });
     
     return {
         isValid: allErrors.length === 0,
         errors: allErrors,
         processedItems: processedItems
     };
 }
 
 /**
  * Calculates totals from processed line items
  * @param {array} processedItems - Array of processed line items
  * @returns {object} - Calculated totals
  */
 static calculateTotals(processedItems) {
     const totalTaxable = processedItems.reduce((sum, item) => sum + (item.line_total || 0), 0);
     const totalTax = processedItems.reduce((sum, item) => sum + (item.tax_amount || 0), 0);
     
     console.log(`📊 LineItemsValidator calculation: ${processedItems.length} items, Taxable: ₹${totalTaxable}, Tax: ₹${totalTax}`);
     
     return {
         totalTaxable: this.roundToTwoDecimals(totalTaxable),
         totalTax: this.roundToTwoDecimals(totalTax),
         grandTotal: this.roundToTwoDecimals(totalTaxable + totalTax)
     };
 }
 
 /**
  * Rounds number to two decimal places
  * @param {number} value - Number to round
  * @returns {number} - Rounded number
  */
 static roundToTwoDecimals(value) {
     return Math.round((value + Number.EPSILON) * 100) / 100;
 }
}

// 🗄️ ENHANCED DATABASE SETUP WITH PARSEUR DOCUMENT ID SUPPORT
function setupDatabase() {
 console.log('🗄️  Setting up database...');
 
 // Check if database exists
 const dbExists = fs.existsSync('./simplifier.db');
 
 if (dbExists) {
     console.log('📁 Found existing database, connecting...');
 } else {
     console.log('🆕 Creating new database...');
 }
 
 db = new sqlite3.Database('./simplifier.db', (err) => {
     if (err) {
         console.error('❌ Database connection failed:', err.message);
         process.exit(1);
     }
     
     if (dbExists) {
         console.log('✅ Connected to existing database');
         
         // Verify table structure and add missing columns if needed
         db.run(`ALTER TABLE expenses ADD COLUMN parseur_document_id TEXT DEFAULT NULL`, (err) => {
             if (err && !err.message.includes('duplicate column name')) {
                 console.log('⚠️ Note: parseur_document_id column may already exist');
             }
         });
         
         db.run(`ALTER TABLE expenses ADD COLUMN file_hash TEXT DEFAULT ""`, (err) => {
             if (err && !err.message.includes('duplicate column name')) {
                 console.log('⚠️ Note: file_hash column may already exist');
             }
         });
         
         // Add other missing columns if needed
         db.run(`ALTER TABLE expenses ADD COLUMN entry_type TEXT DEFAULT "upload"`, (err) => {
             if (err && !err.message.includes('duplicate column name')) {
                 console.log('⚠️ Note: entry_type column may already exist');
             }
         });
         
     } else {
         console.log('✅ Database connected, creating tables...');
         
         // Create the table only if database is new
         const sql = `CREATE TABLE expenses (
             id INTEGER PRIMARY KEY AUTOINCREMENT,
             invoice_date TEXT DEFAULT "",
             invoice_number TEXT DEFAULT "",
             vendor_name TEXT DEFAULT "Unknown Vendor",
             vendor_gstn TEXT DEFAULT "",
             taxable_amount REAL DEFAULT 0,
             igst_amount REAL DEFAULT 0,
             cgst_amount REAL DEFAULT 0,
             sgst_amount REAL DEFAULT 0,
             cess_amount REAL DEFAULT 0,
             round_off REAL DEFAULT 0,
             invoice_value REAL DEFAULT 0,
             tds_rate REAL DEFAULT 0,
             tds_amount REAL DEFAULT 0,
             description TEXT DEFAULT "",
             hsn_sac_code TEXT DEFAULT "",
             line_item_amount REAL DEFAULT 0,
             quantity REAL DEFAULT 0,
             unit_rate REAL DEFAULT 0,
             voucher_type TEXT DEFAULT "Purchase",
             ledger_name TEXT DEFAULT "",
             vendor_address TEXT DEFAULT "",
             customer_gstn TEXT DEFAULT "",
             customer_name TEXT DEFAULT "",
             place_of_supply TEXT DEFAULT "",
             total_tax_amount REAL DEFAULT 0,
             cgst_rate REAL DEFAULT 0,
             sgst_rate REAL DEFAULT 0,
             igst_rate REAL DEFAULT 0,
             cess_rate REAL DEFAULT 0,
             category TEXT DEFAULT "General",
             file_path TEXT DEFAULT "",
             original_filename TEXT DEFAULT "",
             file_hash TEXT DEFAULT "",
             parseur_document_id TEXT DEFAULT NULL,
             extracted_text TEXT DEFAULT "",
             confidence_score REAL DEFAULT 0,
             document_type TEXT DEFAULT "image",
             processing_time_ms INTEGER DEFAULT 0,
             status TEXT DEFAULT "pending_review",
             amount_confidence REAL DEFAULT 0,
             line_items TEXT DEFAULT "[]",
             line_items_count INTEGER DEFAULT 0,
             has_line_items BOOLEAN DEFAULT 0,
             table_structure_confidence REAL DEFAULT 0,
             processing_method TEXT DEFAULT "auto",
             processing_source TEXT DEFAULT "unknown",
             validation_errors TEXT DEFAULT "[]",
             validation_warnings TEXT DEFAULT "[]",
             entry_type TEXT DEFAULT "upload",
             created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
             UNIQUE(invoice_number, vendor_gstn, invoice_value)
         )`;
         
         db.run(sql, (err) => {
             if (err) {
                 console.error('❌ Table creation failed:', err.message);
                 process.exit(1);
             } else {
                 console.log('✅ New database table created successfully');
             }
         });
     }
     
     console.log('✅ Database setup complete');
     app.set('db', db);
     initTallySchema(db).catch(e => console.error('Tally schema init failed:', e.message));
 });
}

// GST State Code Mapping
const GST_STATE_CODES = {
 '01': 'Jammu and Kashmir', '02': 'Himachal Pradesh', '03': 'Punjab', '04': 'Chandigarh',
 '05': 'Uttarakhand', '06': 'Haryana', '07': 'Delhi', '08': 'Rajasthan',
 '09': 'Uttar Pradesh', '10': 'Bihar', '11': 'Sikkim', '12': 'Arunachal Pradesh',
 '13': 'Nagaland', '14': 'Manipur', '15': 'Mizoram', '16': 'Tripura',
 '17': 'Meghalaya', '18': 'Assam', '19': 'West Bengal', '20': 'Jharkhand',
 '21': 'Odisha', '22': 'Chhattisgarh', '23': 'Madhya Pradesh', '24': 'Gujarat',
 '25': 'Daman and Diu', '26': 'Dadra and Nagar Haveli', '27': 'Maharashtra', '28': 'Andhra Pradesh',
 '29': 'Karnataka', '30': 'Goa', '31': 'Lakshadweep', '32': 'Kerala',
 '33': 'Tamil Nadu', '34': 'Puducherry', '35': 'Andaman and Nicobar Islands', '36': 'Telangana',
 '37': 'Andhra Pradesh', '38': 'Ladakh'
};

// Enhanced Line Items Extraction Engine
class LineItemsExtractor {
 static extractLineItemsFromText(text) {
     const lineItems = [];
     const lines = text.split('\n').map(line => line.trim()).filter(line => line.length > 0);
     
     try {
         // Pattern 1: Table-based detection
         const tableItems = this.extractTableBasedItems(lines);
         if (tableItems.length > 0) {
             console.log('✅ Table-based line items detected:', tableItems.length);
             return tableItems;
         }
         
         // Pattern 2: Structured list detection
         const listItems = this.extractStructuredListItems(lines);
         if (listItems.length > 0) {
             console.log('✅ Structured list items detected:', listItems.length);
             return listItems;
         }
         
         // Pattern 3: Invoice section detection
         const sectionItems = this.extractSectionBasedItems(lines);
         if (sectionItems.length > 0) {
             console.log('✅ Section-based items detected:', sectionItems.length);
             return sectionItems;
         }
         
         console.log('⚠️ No line items pattern detected');
         return [];
         
     } catch (error) {
         console.error('❌ Line items extraction failed:', error);
         return [];
     }
 }
 
 static extractTableBasedItems(lines) {
     const items = [];
     let inTableSection = false;
     let tableHeaders = [];
     
     const tableKeywords = [
         'description', 'item', 'product', 'service', 'qty', 'quantity', 'rate', 'amount', 
         'price', 'total', 'hsn', 'sac', 'tax', 'gst', 'unit', 'per'
     ];
     
     for (let i = 0; i < lines.length; i++) {
         const line = lines[i].toLowerCase();
         
         // Detect table headers
         const keywordCount = tableKeywords.filter(keyword => line.includes(keyword)).length;
         if (keywordCount >= 3) {
             inTableSection = true;
             tableHeaders = this.parseTableHeaders(lines[i]);
             console.log('📊 Table headers detected:', tableHeaders);
             continue;
         }
         
         // Extract table rows
         if (inTableSection) {
             const item = this.parseTableRow(lines[i], tableHeaders);
             if (item && item.description) {
                 items.push(item);
             }
             
             // Stop at table end indicators
             if (line.includes('subtotal') || line.includes('total') || line.includes('tax') || 
                 line.includes('discount') || line.includes('net amount') || line.length < 10) {
                 if (items.length > 0) break;
             }
         }
     }
     
     return items;
 }
 
 static extractStructuredListItems(lines) {
     const items = [];
     let currentItem = null;
     
     for (let i = 0; i < lines.length; i++) {
         const line = lines[i];
         const cleanLine = line.replace(/[^\w\s\.\-₹]/g, ' ').trim();
         
         // Skip obvious non-item lines
         if (this.isNonItemLine(cleanLine)) continue;
         
         // Detect item start patterns
         if (this.isItemStartLine(cleanLine)) {
             if (currentItem && currentItem.description) {
                 items.push(currentItem);
             }
             currentItem = this.parseItemLine(line);
         } else if (currentItem && this.isItemContinuation(cleanLine)) {
             this.enhanceItemWithLine(currentItem, line);
         }
     }
     
     // Add final item
     if (currentItem && currentItem.description) {
         items.push(currentItem);
     }
     
     return items.filter(item => item.description && item.description.length > 3);
 }
 
 static extractSectionBasedItems(lines) {
     const items = [];
     let inItemsSection = false;
     
     const sectionKeywords = ['items', 'products', 'services', 'details', 'particulars'];
     
     for (let i = 0; i < lines.length; i++) {
         const line = lines[i].toLowerCase();
         
         // Detect items section start
         if (!inItemsSection && sectionKeywords.some(keyword => line.includes(keyword))) {
             inItemsSection = true;
             continue;
         }
         
         if (inItemsSection) {
             const item = this.parseGenericItemLine(lines[i]);
             if (item && item.description) {
                 items.push(item);
             }
             
             // Stop at section end
             if (line.includes('total') || line.includes('subtotal') || line.includes('tax')) {
                 break;
             }
         }
     }
     
     return items;
 }
 
 static parseTableHeaders(headerLine) {
     const headers = [];
     const parts = headerLine.toLowerCase().split(/[\s\|]+/).filter(p => p.length > 0);
     
     const headerMap = {
         'description': 'description', 'item': 'description', 'product': 'description', 'service': 'description',
         'qty': 'quantity', 'quantity': 'quantity', 'qnt': 'quantity',
         'rate': 'unit_rate', 'price': 'unit_rate', 'unit': 'unit_rate',
         'amount': 'line_total', 'total': 'line_total',
         'hsn': 'hsn_code', 'sac': 'hsn_code',
         'tax': 'tax_rate', 'gst': 'tax_rate'
     };
     
     parts.forEach(part => {
         const mapped = headerMap[part] || part;
         headers.push(mapped);
     });
     
     return headers;
 }
 
 static parseTableRow(rowLine, headers) {
     const item = {
         description: '',
         hsn_code: '',
         quantity: 1,
         unit_rate: 0,
         line_total: 0,
         tax_rate: 18,
         tax_amount: 0
     };
     
     // Split row into columns (handle various separators)
     const columns = rowLine.split(/[\|\t]/).map(col => col.trim());
     if (columns.length === 1) {
         // Try space-based splitting for amounts
         const parts = rowLine.split(/\s+/);
         const amounts = parts.filter(part => /[\d,₹]/.test(part));
         const textParts = parts.filter(part => !/^[\d,₹\.\s]+$/.test(part));
         
         item.description = textParts.join(' ');
         if (amounts.length >= 2) {
             item.quantity = this.parseNumber(amounts[0]) || 1;
             item.unit_rate = this.parseNumber(amounts[amounts.length - 2]) || 0;
             item.line_total = this.parseNumber(amounts[amounts.length - 1]) || 0;
         }
     } else {
         // Map columns to item properties
         headers.forEach((header, index) => {
             if (columns[index]) {
                 const value = columns[index].trim();
                 switch (header) {
                     case 'description':
                         item.description = value;
                         break;
                     case 'quantity':
                         item.quantity = this.parseNumber(value) || 1;
                         break;
                     case 'unit_rate':
                         item.unit_rate = this.parseNumber(value) || 0;
                         break;
                     case 'line_total':
                         item.line_total = this.parseNumber(value) || 0;
                         break;
                     case 'hsn_code':
                         item.hsn_code = value;
                         break;
                     case 'tax_rate':
                         item.tax_rate = this.parseNumber(value) || 18;
                         break;
                 }
             }
         });
     }
     
     // Calculate missing values
     if (item.line_total === 0 && item.quantity > 0 && item.unit_rate > 0) {
         item.line_total = item.quantity * item.unit_rate;
     }
     
     if (item.tax_amount === 0 && item.line_total > 0) {
         item.tax_amount = (item.line_total * item.tax_rate) / 100;
     }
     
     return item.description ? item : null;
 }
 
 static parseItemLine(line) {
     const item = {
         description: '',
         hsn_code: '',
         quantity: 1,
         unit_rate: 0,
         line_total: 0,
         tax_rate: 18,
         tax_amount: 0
     };
     
     // Extract amounts (₹ symbols, decimal numbers)
     const amounts = [...line.matchAll(/₹?\s*([0-9,]+(?:\.[0-9]{2})?)/g)]
         .map(match => this.parseNumber(match[1]))
         .filter(num => num > 0);
     
     // Extract HSN/SAC code
     const hsnMatch = line.match(/\b(\d{4,8})\b/);
     if (hsnMatch) {
         item.hsn_code = hsnMatch[1];
     }
     
     // Extract description (remove amounts and codes)
     item.description = line
         .replace(/₹?\s*[0-9,]+(?:\.[0-9]{2})?/g, '')
         .replace(/\b\d{4,8}\b/g, '')
         .replace(/\s+/g, ' ')
         .trim();
     
     // Assign amounts based on count
     if (amounts.length >= 3) {
         item.quantity = amounts[0];
         item.unit_rate = amounts[1];
         item.line_total = amounts[2];
     } else if (amounts.length === 2) {
         item.unit_rate = amounts[0];
         item.line_total = amounts[1];
     } else if (amounts.length === 1) {
         item.line_total = amounts[0];
     }
     
     // Calculate tax
     if (item.line_total > 0) {
         item.tax_amount = (item.line_total * item.tax_rate) / 100;
     }
     
     return item;
 }
 
 static parseGenericItemLine(line) {
     // Simple fallback parser
     const item = {
         description: line.replace(/[₹\d,\.]/g, '').trim(),
         quantity: 1,
         unit_rate: 0,
         line_total: 0,
         tax_rate: 18,
         tax_amount: 0
     };
     
     const amounts = [...line.matchAll(/₹?\s*([0-9,]+(?:\.[0-9]{2})?)/g)]
         .map(match => this.parseNumber(match[1]));
     
     if (amounts.length > 0) {
         item.line_total = amounts[amounts.length - 1];
         item.tax_amount = (item.line_total * 18) / 100;
     }
     
     return item.description.length > 3 ? item : null;
 }
 
 static enhanceItemWithLine(item, line) {
     // Add continuation text to description
     const cleanLine = line.replace(/[₹\d,\.]/g, '').trim();
     if (cleanLine.length > 3 && !item.description.includes(cleanLine)) {
         item.description += ' ' + cleanLine;
     }
     
     // Extract additional amounts
     const amounts = [...line.matchAll(/₹?\s*([0-9,]+(?:\.[0-9]{2})?)/g)]
         .map(match => this.parseNumber(match[1]));
     
     if (amounts.length > 0 && item.line_total === 0) {
         item.line_total = amounts[amounts.length - 1];
     }
 }
 
 static isNonItemLine(line) {
     const nonItemKeywords = [
         'invoice', 'bill', 'total', 'subtotal', 'tax', 'discount', 'payment', 'due',
         'address', 'phone', 'email', 'website', 'gstin', 'terms', 'conditions',
         'thank', 'regards', 'signature', 'stamp', 'page', 'continued'
     ];
     
     const lineLower = line.toLowerCase();
     return nonItemKeywords.some(keyword => lineLower.includes(keyword)) || 
            line.length < 5 || 
            /^[\d\s\-\.\/]+$/.test(line);
 }
 
 static isItemStartLine(line) {
     return line.length > 10 && 
            /[a-zA-Z]/.test(line) && 
            /\d/.test(line) && !this.isNonItemLine(line);
 }
 
 static isItemContinuation(line) {
     return line.length > 3 && 
            /[a-zA-Z]/.test(line) && 
            !this.isItemStartLine(line);
 }
 
 static parseNumber(str) {
     if (!str) return 0;
     const cleaned = str.toString().replace(/[₹,\s]/g, '');
     const num = parseFloat(cleaned);
     return isNaN(num) ? 0 : Math.round(num * 100) / 100;
 }
 
 static calculateTableConfidence(lineItems) {
     if (lineItems.length === 0) return 0;
     
     let score = 0.5;
     
     const completeItems = lineItems.filter(item => 
         item.description && item.quantity > 0 && item.line_total > 0
     );
     score += (completeItems.length / lineItems.length) * 0.3;
     
     const itemsWithHSN = lineItems.filter(item => item.hsn_code);
     score += (itemsWithHSN.length / lineItems.length) * 0.2;
     
     return Math.min(score, 0.99);
 }
}

// Enhanced AI Line Items Processor
class AILineItemsProcessor {
 static async enhanceLineItemsWithGPT(rawText, extractedLineItems) {
     try {
         if (!process.env.OPENAI_API_KEY) {
             console.log('⚠️ OpenAI API key not available for line items enhancement');
             return extractedLineItems;
         }
         
         if (extractedLineItems.length === 0) {
             console.log('🤖 Using GPT to extract line items from text...');
             return await this.extractLineItemsWithGPT(rawText);
         } else {
             console.log('🤖 Using GPT to enhance existing line items...');
             return await this.enhanceWithGPT(rawText, extractedLineItems);
         }
         
     } catch (error) {
         console.error('❌ GPT line items processing failed:', error);
         return extractedLineItems;
     }
 }
 
 static async extractLineItemsWithGPT(rawText) {
     try {
         const response = await openai.chat.completions.create({
             model: "gpt-4o-mini",
             messages: [
                 {
                     role: "system",
                     content: `You are an expert at extracting line items from invoice text. Extract all line items with the following structure:

{
"line_items": [
{
  "description": "item description",
  "hsn_code": "HSN/SAC code if available",
  "quantity": number,
  "unit_rate": number,
  "line_total": number,
  "tax_rate": number (default 18 if not specified),
  "tax_amount": number
}
]
}

Rules:
- Extract ALL items/products/services mentioned
- Calculate missing tax_amount as (line_total * tax_rate / 100)
- Default quantity to 1 if not specified
- Include HSN/SAC codes when available
- Return empty array if no line items found
- Return ONLY valid JSON`
                 },
                 {
                     role: "user",
                     content: `Extract line items from this invoice text:\n\n${rawText}`
                 }
             ],
             temperature: 0.1,
             max_tokens: 1500
         });
         
         const gptResponse = response.choices[0].message.content.trim();
         const parsed = JSON.parse(gptResponse.replace(/```json|```/g, ''));
         
         console.log(`✅ GPT extracted ${parsed.line_items?.length || 0} line items`);
         return parsed.line_items || [];
         
     } catch (error) {
         console.error('❌ GPT line items extraction failed:', error);
         return [];
     }
 }
 
 static async enhanceWithGPT(rawText, lineItems) {
     try {
         const response = await openai.chat.completions.create({
             model: "gpt-4o-mini",
             messages: [
                 {
                     role: "system",
                     content: `Enhance and correct the extracted line items using the full invoice text. Fix any missing data, correct amounts, and ensure accuracy. Return the same JSON structure with enhanced data.`
                 },
                 {
                     role: "user",
                     content: `Original Invoice Text:\n${rawText}\n\nExtracted Line Items:\n${JSON.stringify(lineItems, null, 2)}\n\nEnhance and return corrected line items as JSON array.`
                 }
             ],
             temperature: 0.1,
             max_tokens: 1500
         });
         
         const enhanced = JSON.parse(response.choices[0].message.content.trim().replace(/```json|```/g, ''));
         console.log(`✅ GPT enhanced ${enhanced.length || 0} line items`);
         return enhanced;
         
     } catch (error) {
         console.error('❌ GPT line items enhancement failed:', error);
         return lineItems;
     }
 }
}

// 🧠 ENHANCED PARSEUR API HELPER CLASS WITH REPROCESSING SUPPORT
class ParseurInvoiceProcessor {
 constructor(config) {
     this.config = config;
     this.headers = {
         'Authorization': `Token ${config.apiToken}`,
         'Content-Type': 'application/json'
     };
 }

 async testConnection() {
     try {
         const response = await axios.get(
             `https://api.parseur.com/parser/${this.config.mailboxId}`,
             {
                 headers: {
                     'Authorization': `Token ${this.config.apiToken}`
                 },
                 timeout: 10000
             }
         );
         return { success: true, data: response.data };
     } catch (error) {
         return { success: false, error: error.response?.data || error.message };
     }
 }

 /**
  * 🧠 REPROCESSING LOGIC: Reuse existing Parseur document ID instead of uploading again
  * @param {string} documentId - Existing Parseur document ID
  * @returns {Promise<object>} - Reprocessing result
  */
 async reprocessDocument(documentId) {
     try {
         console.log('🔄 Reprocessing existing Parseur document:', documentId);
         
         // Get processed results using existing document ID
         const processResult = await this.getProcessedResults(documentId, 5, 2000); // Shorter retry for existing docs
         
         if (!processResult.success) {
             throw new Error(`Reprocessing failed: ${processResult.error}`);
         }

         console.log('✅ Document reprocessed successfully from existing Parseur data');

         return {
             success: true,
             extractedData: {
                 ...processResult.extractedData,
                 processing_time_ms: 500, // Fast reprocessing
                 confidence_score: processResult.confidence,
                 processing_source: 'parseur_reprocessed'
             },
             rawText: processResult.rawText,
             processingTime: 500,
             documentId: documentId,
             reprocessed: true
         };

     } catch (error) {
         console.error('❌ Parseur reprocessing failed:', error);
         throw error;
     }
 }

 getContentType(filePath) {
     const ext = path.extname(filePath).toLowerCase();
     const contentTypes = {
         '.pdf': 'application/pdf',
         '.jpg': 'image/jpeg',
         '.jpeg': 'image/jpeg',
         '.png': 'image/png',
         '.gif': 'image/gif',
         '.bmp': 'image/bmp',
         '.tiff': 'image/tiff',
         '.webp': 'image/webp'
     };
     return contentTypes[ext] || 'application/octet-stream';
 }

 async sendInvoiceDocument(filePath, originalName) {
     try {
         const connectionTest = await this.testConnection();
         if (!connectionTest.success) {
             console.log('⚠️ Connection test failed, proceeding with upload attempt...');
         }
         
         try {
             const formData = new FormData();
             formData.append('file', fs.createReadStream(filePath), {
                 filename: originalName,
                 contentType: this.getContentType(filePath)
             });
             
             const response = await axios.post(
                 `https://api.parseur.com/parser/${this.config.mailboxId}/upload`,
                 formData,
                 {
                     headers: {
                         ...formData.getHeaders(),
                         'Authorization': `Token ${this.config.apiToken}`
                     },
                     timeout: 60000
                 }
             );

             let documentId = null;
             if (response.data.attachments && response.data.attachments.length > 0) {
                 documentId = response.data.attachments[0].DocumentID;
             } else if (response.data.id) {
                 documentId = response.data.id;
             } else if (response.data.document_id) {
                 documentId = response.data.document_id;
             }

             if (!documentId) {
                 throw new Error('No DocumentID found in upload response');
             }

             return {
                 success: true,
                 documentId: documentId,
                 status: response.data.status || 'uploaded',
                 uploadResponse: response.data
             };
             
         } catch (method1Error) {
             return await this.sendInvoiceDocumentBase64(filePath, originalName);
         }
         
     } catch (error) {
         return {
             success: false,
             error: error.response?.data || error.message
         };
     }
 }

 async sendInvoiceDocumentBase64(filePath, originalName) {
     try {
         const fileBuffer = fs.readFileSync(filePath);
         const base64Data = fileBuffer.toString('base64');
         
         const payload = {
             document: {
                 name: originalName,
                 content: base64Data,
                 content_type: this.getContentType(filePath)
             }
         };
         
         const response = await axios.post(
             `https://api.parseur.com/parser/${this.config.mailboxId}/documents`,
             payload,
             {
                 headers: {
                     'Authorization': `Token ${this.config.apiToken}`,
                     'Content-Type': 'application/json'
                 },
                 timeout: 60000
             }
         );

         let documentId = null;
         if (response.data.attachments && response.data.attachments.length > 0) {
             documentId = response.data.attachments[0].DocumentID;
         } else if (response.data.id) {
             documentId = response.data.id;
         } else if (response.data.document_id) {
             documentId = response.data.document_id;
         }

         if (!documentId) {
             throw new Error('No DocumentID found in upload response');
         }

         return {
             success: true,
             documentId: documentId,
             status: response.data.status || 'uploaded',
             uploadResponse: response.data
         };
         
     } catch (error) {
         return {
             success: false,
             error: error.response?.data || error.message
         };
     }
 }

 async getProcessedResults(documentId, maxRetries = 10, retryDelay = 3000) {
     try {
         if (!documentId || documentId === 'undefined' || documentId === null) {
             throw new Error('Invalid DocumentID provided: ' + documentId);
         }
         
         for (let attempt = 1; attempt <= maxRetries; attempt++) {
             try {
                 const endpoints = [
                     `https://api.parseur.com/document/${documentId}`,
                     `https://api.parseur.com/parser/${this.config.mailboxId}/document/${documentId}`,
                     `https://api.parseur.com/documents/${documentId}`,
                     `https://api.parseur.com/parser/${this.config.mailboxId}/documents/${documentId}`,
                     `https://api.parseur.com/parser/${this.config.mailboxId}/document_set/${documentId}`
                 ];
                 
                 let response = null;
                 
                 for (const endpoint of endpoints) {
                     try {
                         response = await axios.get(endpoint, {
                             headers: {
                                 'Authorization': `Token ${this.config.apiToken}`
                             },
                             timeout: 10000
                         });
                         break;
                     } catch (endpointError) {
                         continue;
                     }
                 }
                 
                 if (!response) {
                     throw new Error('All endpoints failed');
                 }

                 const document = response.data;
                 const documentData = document.results ? document.results[0] : document;
                 
                 if (documentData.status === 'PARSEDOK') {
                     const parsedData = documentData.parsed_data || documentData.data || documentData;
                     
                     return {
                         success: true,
                         extractedData: this.convertParseurToInternalFormat(parsedData),
                         rawText: documentData.text || '',
                         confidence: documentData.confidence || 0.95,
                         documentId: documentId,
                         fullResponse: documentData
                     };
                 } else if (documentData.status === 'PARSEDKO' || documentData.status === 'INVALID') {
                     return {
                         success: false,
                         error: `Processing failed with status: ${documentData.status}`
                     };
                 } else if (documentData.status === 'QUOTAEXC') {
                     return {
                         success: false,
                         error: 'Quota exceeded'
                     };
                 }

                 if (attempt < maxRetries) {
                     await new Promise(resolve => setTimeout(resolve, retryDelay));
                 }

             } catch (requestError) {
                 if (attempt === maxRetries) {
                     throw requestError;
                 }
                 await new Promise(resolve => setTimeout(resolve, retryDelay));
             }
         }

         return {
             success: false,
             error: 'Processing timeout'
         };

     } catch (error) {
         return {
             success: false,
             error: error.message
         };
     }
 }

 convertParseurToInternalFormat(parseurData) {
     if (!parseurData) {
         console.log('⚠️ No parseur data received');
         return {};
     }

     let actualInvoiceData = {};
     
     if (parseurData.result) {
         try {
             actualInvoiceData = JSON.parse(parseurData.result);
             console.log('📊 Parsed result field successfully');
         } catch (error) {
             console.log('⚠️ Failed to parse result field, using raw data');
             actualInvoiceData = parseurData;
         }
     } else {
         actualInvoiceData = parseurData;
     }

     console.log('🔍 Processing Parseur data with exact field mappings:', Object.keys(actualInvoiceData));

     const extractedData = {
         invoice_date: '',
         invoice_number: '',
         vendor_name: '',
         vendor_gstn: '',
         customer_name: '',
         customer_gstn: '',
         place_of_supply: '',
         taxable_amount: 0,
         igst_amount: 0,
         cgst_amount: 0,
         sgst_amount: 0,
         cess_amount: 0,
         round_off: 0,
         invoice_value: 0,
         tds_rate: 0,
         tds_amount: 0,
         description: '',
         hsn_sac_code: '',
         line_item_amount: 0,
         quantity: 0,
         unit_rate: 0,
         voucher_type: 'Purchase',
         ledger_name: '',
         vendor_address: '',
         total_tax_amount: 0,
         cgst_rate: 0,
         sgst_rate: 0,
         igst_rate: 0,
         cess_rate: 0,
         category: 'General',
         line_items: []
     };

     const fieldMappings = {
         invoice_date: [
             'Invoice Date', 'invoice_date', 'Invoice_Date', 'date', 'Date',
             'Bill Date', 'bill_date', 'Document Date', 'document_date'
         ],
         invoice_number: [
             'Invoice Number', 'invoice_number', 'Invoice No', 'invoice_no',
             'Bill Number', 'bill_number', 'Document Number', 'document_number',
             'Invoice#', 'invoice#', 'Inv No', 'inv_no'
         ],
         vendor_name: [
             'Vendor Name', 'vendor_name', 'Supplier Name', 'supplier_name',
             'Company Name', 'company_name', 'Business Name', 'business_name',
             'Vendor', 'Supplier', 'From', 'Seller'
         ],
         vendor_gstn: [
             'Vendor GSTN', 'vendor_gstn', 'Supplier GSTN', 'supplier_gstn',
             'GSTIN', 'gstin', 'GST Number', 'gst_number', 'Tax ID', 'tax_id',
             'Vendor GST', 'vendor_gst', 'Supplier GST', 'supplier_gst'
         ],
         customer_name: [
             'Customer Name', 'customer_name', 'Buyer Name', 'buyer_name',
             'Recipient Name', 'recipient_name', 'To Name', 'to_name',
             'Client Name', 'client_name'
         ],
         customer_gstn: [
             'Customer GSTN', 'customer_gstn', 'Buyer GSTN', 'buyer_gstn',
             'Customer GST', 'customer_gst', 'Buyer GST', 'buyer_gst',
             'Recipient GSTN', 'recipient_gstn', 'To GSTN', 'to_gstn'
         ],
         place_of_supply: [
             'Place_of_Supply', 'place_of_supply', 'Place of Supply', 'place_supply',
             'Supply Place', 'supply_place'
         ],
         taxable_amount: [
             'Taxable Amount', 'taxable_amount', 'Taxable Value', 'taxable_value',
             'Base Amount', 'base_amount', 'Net Amount', 'net_amount',
             'Subtotal', 'subtotal', 'Sub Total', 'sub_total'
         ],
         igst_amount: [
             'IGST amount', 'igst_amount', 'IGST Amount', 'IGST_amount',
             'IGST', 'igst', 'Integrated GST', 'integrated_gst'
         ],
         cgst_amount: [
             'CGST amount', 'cgst_amount', 'CGST Amount', 'CGST_amount',
             'CGST', 'cgst', 'Central GST', 'central_gst'
         ],
         sgst_amount: [
             'SGST amount', 'sgst_amount', 'SGST Amount', 'SGST_amount',
             'SGST', 'sgst', 'State GST', 'state_gst'
         ],
         invoice_value: [
             'Total Invoice Value', 'total_invoice_value', 'Invoice Value', 'invoice_value',
             'Total Amount', 'total_amount', 'Grand Total', 'grand_total',
             'Final Amount', 'final_amount', 'Total', 'total'
         ],
         cess_amount: [
             'CESS amount', 'cess_amount', 'Cess amount', 'Cess Amount',
             'CESS', 'cess', 'Cess', 'Education Cess', 'education_cess'
         ],
         round_off: [
             'Round Off', 'round_off', 'Rounding', 'rounding',
             'Round Off Amount', 'round_off_amount'
         ],
         tds_rate: [
             'TDS Rate', 'tds_rate', 'TDS%', 'tds_percentage'
         ],
         tds_amount: [
             'TDS amount', 'tds_amount', 'TDS Amount', 'TDS',
             'Tax Deducted', 'tax_deducted'
         ],
         description: [
             'Description', 'description', 'Item Description', 'item_description',
             'Product Description', 'product_description', 'Details', 'details'
         ],
         hsn_sac_code: [
             'HSN/SAC Code', 'hsn_sac_code', 'hsnSacCode', 'HSN Code', 'hsn_code',
             'SAC Code', 'sac_code', 'HSN', 'hsn', 'SAC', 'sac'
         ]
     };

     for (const [internalField, possibleNames] of Object.entries(fieldMappings)) {
         let foundValue = null;
         
         for (const fieldName of possibleNames) {
             if (actualInvoiceData.hasOwnProperty(fieldName)) {
                 const value = actualInvoiceData[fieldName];
                 if (value !== undefined && value !== '') {
                     foundValue = value;
                     break;
                 }
             }
             
             const keys = Object.keys(actualInvoiceData);
             const matchedKey = keys.find(key => key.toLowerCase() === fieldName.toLowerCase());
             if (matchedKey) {
                 const value = actualInvoiceData[matchedKey];
                 if (value !== undefined && value !== '') {
                     foundValue = value;
                     break;
                 }
             }
         }
         
         if (foundValue !== null) {
             if (internalField.includes('amount') || internalField.includes('value') || 
                 internalField === 'quantity' || internalField === 'unit_rate') {
                 extractedData[internalField] = foundValue === null ? 0 : this.parseAmount(foundValue);
             } else if (internalField.includes('rate') && internalField !== 'unit_rate') {
                 extractedData[internalField] = foundValue === null ? 0 : (parseFloat(foundValue) || 0);
             } else if (internalField.includes('gstn')) {
                 extractedData[internalField] = foundValue === null ? '' : this.cleanGSTN(foundValue);
             } else if (internalField === 'invoice_date') {
                 extractedData[internalField] = foundValue === null ? '' : this.parseDate(foundValue);
             } else {
                 extractedData[internalField] = foundValue === null ? '' : foundValue.toString();
             }
             
             console.log(`✅ Mapped ${internalField}: ${foundValue}`);
         }
     }

     const lineItemsFieldNames = [
         'items', 'line_items', 'lineItems', 'invoice_items', 'products',
         'services', 'details', 'item_details', 'product_details',
         'line_item_details', 'table_data', 'item_list'
     ];
     
     let extractedLineItems = [];
     
     for (const fieldName of lineItemsFieldNames) {
         if (actualInvoiceData[fieldName] && Array.isArray(actualInvoiceData[fieldName])) {
             console.log(`📊 Found line items in field: ${fieldName}`);
             extractedLineItems = this.processLineItemsArray(actualInvoiceData[fieldName]);
             break;
         }
         
         const keys = Object.keys(actualInvoiceData);
         const matchedKey = keys.find(key => key.toLowerCase() === fieldName.toLowerCase());
         if (matchedKey && actualInvoiceData[matchedKey] && Array.isArray(actualInvoiceData[matchedKey])) {
             console.log(`📊 Found line items in field: ${matchedKey}`);
             extractedLineItems = this.processLineItemsArray(actualInvoiceData[matchedKey]);
             break;
         }
     }

     if (extractedLineItems.length > 0) {
         extractedData.line_items = extractedLineItems;
         console.log(`✅ Processed ${extractedLineItems.length} line items from Parseur`);
         
         // 🆕 CALCULATE TOTALS FROM LINE ITEMS
         const lineItemTotals = this.calculateLineItemTotals(extractedLineItems);
         
         // Update invoice totals if line items provide better data
         if (lineItemTotals.totalTaxable > 0) {
             extractedData.taxable_amount = lineItemTotals.totalTaxable;
             console.log(`📊 Updated taxable amount from line items: ₹${lineItemTotals.totalTaxable}`);
         }
         
         if (lineItemTotals.totalTax > 0) {
             extractedData.total_tax_amount = lineItemTotals.totalTax;
             console.log(`📊 Updated total tax from line items: ₹${lineItemTotals.totalTax}`);
         }
         
         // Calculate invoice value if we have taxable amount
         if (extractedData.taxable_amount > 0) {
             const calculatedInvoiceValue = extractedData.taxable_amount + 
                                          (extractedData.total_tax_amount || 0) + 
                                          (extractedData.round_off || 0) - 
                                          (extractedData.tds_amount || 0);
             
             // Only update if the calculated value is significantly different or if invoice_value is 0
             if (extractedData.invoice_value === 0 || Math.abs(extractedData.invoice_value - calculatedInvoiceValue) > 1) {
                 extractedData.invoice_value = calculatedInvoiceValue;
                 console.log(`📊 Updated invoice value from line items: ₹${calculatedInvoiceValue}`);
             }
         }
         
         if (extractedLineItems.length === 1) {
             const singleItem = extractedLineItems[0];
             if (!extractedData.description || extractedData.description.length < singleItem.description.length) {
                 extractedData.description = singleItem.description;
             }
             if (!extractedData.hsn_sac_code && singleItem.hsn_code) {
                 extractedData.hsn_sac_code = singleItem.hsn_code;
             }
             extractedData.quantity = singleItem.quantity;
             extractedData.unit_rate = singleItem.unit_rate;
             extractedData.line_item_amount = singleItem.line_total;
         }
     } else {
         console.log('⚠️ No line items found in Parseur response');
     }

     extractedData.total_tax_amount = (extractedData.cgst_amount || 0) + 
                                    (extractedData.sgst_amount || 0) + 
                                    (extractedData.igst_amount || 0) + 
                                    (extractedData.cess_amount || 0);

     if (extractedData.taxable_amount > 0) {
         if (extractedData.cgst_amount > 0) {
             extractedData.cgst_rate = Math.round((extractedData.cgst_amount / extractedData.taxable_amount) * 100 * 100) / 100;
         }
         if (extractedData.sgst_amount > 0) {
             extractedData.sgst_rate = Math.round((extractedData.sgst_amount / extractedData.taxable_amount) * 100 * 100) / 100;
         }
         if (extractedData.igst_amount > 0) {
             extractedData.igst_rate = Math.round((extractedData.igst_amount / extractedData.taxable_amount) * 100 * 100) / 100;
         }
         if (extractedData.cess_amount > 0) {
             extractedData.cess_rate = Math.round((extractedData.cess_amount / extractedData.taxable_amount) * 100 * 100) / 100;
         }
     }

     extractedData.amount = extractedData.invoice_value;
     extractedData.taxable_value = extractedData.taxable_amount;
     
     console.log('✅ Parseur conversion complete with exact field mappings:', {
         vendor: extractedData.vendor_name,
         amount: extractedData.invoice_value,
         lineItems: extractedData.line_items.length,
         voucherType: extractedData.voucher_type
     });
     
     return extractedData;
 }

 processLineItemsArray(itemsArray) {
     const processedItems = [];
     
     itemsArray.forEach((item, index) => {
         if (!item || typeof item !== 'object') return;
         
         const processedItem = {
             description: '',
             hsn_code: '',
             quantity: 1,
             unit_rate: 0,
             line_total: 0,
             tax_rate: 18,
             tax_amount: 0
         };
         
         const lineItemMappings = {
             description: [
                 'Description', 'description', 'Item Description', 'item_description',
                 'Product Description', 'product_description', 'Item', 'item',
                 'Product', 'product', 'Service', 'service', 'Details', 'details',
                 'Product Name', 'product_name', 'Item Name', 'item_name'
             ],
             hsn_code: [
                 'hsnSacCode', 'hsn_sac_code', 'HSN/SAC Code', 'HSN Code', 'hsn_code',
                 'SAC Code', 'sac_code', 'HSN', 'hsn', 'SAC', 'sac',
                 'Product Code', 'product_code', 'Item Code', 'item_code'
             ],
             quantity: [
                 'quantity', 'Quantity', 'qty', 'Qty', 'QTY',
                 'Amount', 'amount', 'Count', 'count', 'Units', 'units'
             ],
             unit_rate: [
                 'unit_rate', 'Unit Rate', 'unitRate', 'unit_price', 'Unit Price',
                 'rate', 'Rate', 'price', 'Price', 'Cost', 'cost',
                 'Unit Cost', 'unit_cost', 'Per Unit', 'per_unit'
             ],
             line_total: [
                 'line_item_taxable amount', 'line_item_amount', 'Line Item Amount',
                 'line_total', 'Line Total', 'lineTotal', 'total', 'Total',
                 'amount', 'Amount', 'Line Amount', 'line_amount',
                 'Taxable Amount', 'taxable_amount', 'Net Amount', 'net_amount',
                 'line_item_invoice amount'
             ],
             tax_rate: [
                 'GST_Percentage', 'gst_percentage', 'GST Percentage', 'tax_rate',
                 'Tax Rate', 'taxRate', 'gst_rate', 'GST Rate', 'GST_Rate',
                 'tax_percent', 'Tax Percent', 'Rate', 'rate'
             ],
             tax_amount: [
                 'GST amount', 'gst_amount', 'GST Amount', 'tax_amount',
                 'Tax Amount', 'taxAmount', 'Tax', 'tax', 'GST', 'gst'
             ]
         };
         
         for (const [field, possibleNames] of Object.entries(lineItemMappings)) {
             let foundValue = null;
             
             for (const fieldName of possibleNames) {
                 if (item.hasOwnProperty(fieldName)) {
                     const value = item[fieldName];
                     if (value !== undefined && value !== '') {
                         foundValue = value;
                         break;
                     }
                 }
                 
                 const keys = Object.keys(item);
                 const matchedKey = keys.find(key => key.toLowerCase() === fieldName.toLowerCase());
                 if (matchedKey) {
                     const value = item[matchedKey];
                     if (value !== undefined && value !== '') {
                         foundValue = value;
                         break;
                     }
                 }
             }
             
             if (foundValue !== null) {
                 if (field === 'description' || field === 'hsn_code') {
                     processedItem[field] = foundValue === null ? '' : foundValue.toString().trim();
                 } else {
                     processedItem[field] = foundValue === null ? 0 : this.parseAmount(foundValue);
                 }
             }
         }
         
         if (processedItem.line_total === 0 && processedItem.quantity > 0 && processedItem.unit_rate > 0) {
             processedItem.line_total = this.roundToTwoDecimals(processedItem.quantity * processedItem.unit_rate);
         }
         
         if (processedItem.tax_amount === 0 && processedItem.line_total > 0 && processedItem.tax_rate > 0) {
             processedItem.tax_amount = this.roundToTwoDecimals((processedItem.line_total * processedItem.tax_rate) / 100);
         }
         
         if (processedItem.description || processedItem.line_total > 0) {
             processedItems.push(processedItem);
             console.log(`✅ Processed Parseur line item ${index + 1}:`, processedItem.description);
         }
     });
     
     return processedItems;
 }

 cleanGSTN(gstn) {
     if (!gstn) return '';
     const cleaned = gstn.replace(/[^A-Z0-9]/g, '').toUpperCase();
     return cleaned.length === 15 ? cleaned : '';
 }

 parseDate(dateStr) {
     return UniversalDateParser.parseInvoiceDate(dateStr);
 }

 parseAmount(amountStr) {
     if (!amountStr) return 0;
     
     let cleaned = amountStr.toString()
         .replace(/[₹,Rs\s]/g, '')
         .replace(/[^\d\.]/g, '');
     
     const decimalParts = cleaned.split('.');
     if (decimalParts.length > 2) {
         cleaned = decimalParts.slice(0, -1).join('') + '.' + decimalParts[decimalParts.length - 1];
     }
     
     const amount = parseFloat(cleaned) || 0;
     return this.roundToTwoDecimals(amount);
 }

 roundToTwoDecimals(value) {
     return Math.round((value + Number.EPSILON) * 100) / 100;
 }

 // 🆕 ADD THIS NEW METHOD
 calculateLineItemTotals(lineItems) {
     if (!lineItems || !Array.isArray(lineItems) || lineItems.length === 0) {
         return { totalTaxable: 0, totalTax: 0, grandTotal: 0 };
     }
     
     let totalTaxable = 0;
     let totalTax = 0;
     
     lineItems.forEach(item => {
         const lineTotal = parseFloat(item.line_total) || 0;
         const taxAmount = parseFloat(item.tax_amount) || 0;
         
         totalTaxable += lineTotal;
         totalTax += taxAmount;
     });
     
     const grandTotal = totalTaxable + totalTax;
     
     console.log(`📊 Line items calculation: Taxable: ₹${totalTaxable}, Tax: ₹${totalTax}, Total: ₹${grandTotal}`);
     
     return {
         totalTaxable: this.roundToTwoDecimals(totalTaxable),
         totalTax: this.roundToTwoDecimals(totalTax),
         grandTotal: this.roundToTwoDecimals(grandTotal)
     };
 }

 async processInvoiceWithParseur(filePath, originalName, voucherType = 'Purchase') {
     try {
         const startTime = Date.now();

         console.log('📤 Sending invoice to Parseur:', originalName);
         const sendResult = await this.sendInvoiceDocument(filePath, originalName);
         
         if (!sendResult.success) {
             throw new Error(`Failed to send to Parseur: ${sendResult.error}`);
         }

         console.log('⏳ Waiting for Parseur processing...');
         const processResult = await this.getProcessedResults(sendResult.documentId);
         
         if (!processResult.success) {
             throw new Error(`Parseur processing failed: ${processResult.error}`);
         }

         const processingTime = Date.now() - startTime;

         processResult.extractedData.voucher_type = voucherType;

         console.log('✅ Parseur processing complete with exact field mappings:', {
             vendor: processResult.extractedData.vendor_name,
             amount: processResult.extractedData.invoice_value,
             lineItems: processResult.extractedData.line_items?.length || 0,
             voucherType: processResult.extractedData.voucher_type
         });

         return {
             success: true,
             extractedData: {
                 ...processResult.extractedData,
                 processing_time_ms: processingTime,
                 confidence_score: processResult.confidence,
                 processing_source: 'parseur',
                 parseur_document_id: sendResult.documentId // 🧠 Store document ID for reprocessing
             },
             rawText: processResult.rawText,
             processingTime: processingTime,
             documentId: sendResult.documentId
         };

     } catch (error) {
         console.error('❌ Parseur processing failed:', error);
         throw error;
     }
 }
}

// 🗺️ GSTIN STATE CODE MAP
const INDIAN_STATE_CODES = {
 '01': 'Jammu & Kashmir', '02': 'Himachal Pradesh', '03': 'Punjab',
 '04': 'Chandigarh', '05': 'Uttarakhand', '06': 'Haryana',
 '07': 'Delhi', '08': 'Rajasthan', '09': 'Uttar Pradesh',
 '10': 'Bihar', '11': 'Sikkim', '12': 'Arunachal Pradesh',
 '13': 'Nagaland', '14': 'Manipur', '15': 'Mizoram',
 '16': 'Tripura', '17': 'Meghalaya', '18': 'Assam',
 '19': 'West Bengal', '20': 'Jharkhand', '21': 'Odisha',
 '22': 'Chhattisgarh', '23': 'Madhya Pradesh', '24': 'Gujarat',
 '26': 'Dadra & Nagar Haveli and Daman & Diu', '27': 'Maharashtra',
 '28': 'Andhra Pradesh', '29': 'Karnataka', '30': 'Goa',
 '31': 'Lakshadweep', '32': 'Kerala', '33': 'Tamil Nadu',
 '34': 'Puducherry', '35': 'Andaman & Nicobar Islands',
 '36': 'Telangana', '37': 'Andhra Pradesh (New)', '38': 'Ladakh',
 '97': 'Other Territory', '99': 'Centre Jurisdiction'
};

// 🔍 GSTIN VALIDATOR
class GSTINValidator {
 static GSTIN_PATTERN = /^[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1}$/;

 static validate(gstin) {
     if (!gstin || typeof gstin !== 'string') {
         return { valid: false, error: 'GSTIN is empty or invalid type' };
     }
     const g = gstin.trim().toUpperCase();
     if (g.length !== 15) {
         return { valid: false, error: `GSTIN must be 15 characters, got ${g.length}` };
     }
     if (!this.GSTIN_PATTERN.test(g)) {
         return { valid: false, error: 'GSTIN format invalid (expected: 2-digit state code + 10-char PAN + entity + Z + checksum)' };
     }
     const stateCode = g.substring(0, 2);
     if (!INDIAN_STATE_CODES[stateCode]) {
         return { valid: false, error: `Unknown GST state code: ${stateCode}` };
     }
     return { valid: true, stateCode, stateName: INDIAN_STATE_CODES[stateCode] };
 }

 // Best-effort: look for state/city keywords in the address string
 static inferStateFromAddress(address) {
     if (!address) return null;
     const addr = address.toLowerCase();
     for (const [, name] of Object.entries(INDIAN_STATE_CODES)) {
         if (addr.includes(name.toLowerCase())) return name;
     }
     const cityMap = {
         'delhi': 'Delhi', 'new delhi': 'Delhi',
         'mumbai': 'Maharashtra', 'pune': 'Maharashtra', 'nagpur': 'Maharashtra',
         'bangalore': 'Karnataka', 'bengaluru': 'Karnataka',
         'chennai': 'Tamil Nadu', 'coimbatore': 'Tamil Nadu',
         'hyderabad': 'Telangana', 'secunderabad': 'Telangana',
         'kolkata': 'West Bengal',
         'ahmedabad': 'Gujarat', 'surat': 'Gujarat', 'vadodara': 'Gujarat',
         'jaipur': 'Rajasthan', 'jodhpur': 'Rajasthan',
         'lucknow': 'Uttar Pradesh', 'noida': 'Uttar Pradesh', 'agra': 'Uttar Pradesh',
         'chandigarh': 'Chandigarh',
         'gurgaon': 'Haryana', 'gurugram': 'Haryana', 'faridabad': 'Haryana',
         'bhopal': 'Madhya Pradesh', 'indore': 'Madhya Pradesh',
         'patna': 'Bihar',
         'bhubaneswar': 'Odisha',
         'raipur': 'Chhattisgarh',
         'guwahati': 'Assam',
         'ranchi': 'Jharkhand',
         'dehradun': 'Uttarakhand',
         'thiruvananthapuram': 'Kerala', 'kochi': 'Kerala',
         'visakhapatnam': 'Andhra Pradesh (New)'
     };
     for (const [city, state] of Object.entries(cityMap)) {
         if (addr.includes(city)) return state;
     }
     return null;
 }
}

// 🤖 CLAUDE AI INVOICE PROCESSOR — PRIMARY EXTRACTOR
class ClaudeInvoiceProcessor {
 constructor(config) {
     this.config = config;
     this.client = anthropic;
 }

 getMediaType(filePath) {
     const ext = path.extname(filePath).toLowerCase();
     const types = {
         '.jpg': 'image/jpeg', '.jpeg': 'image/jpeg',
         '.png': 'image/png', '.gif': 'image/gif',
         '.webp': 'image/webp', '.bmp': 'image/jpeg',
         '.tiff': 'image/jpeg', '.pdf': 'application/pdf'
     };
     return types[ext] || 'image/jpeg';
 }

 // All required fields must be present for extraction to be considered complete
 isExtractionComplete(data) {
     if (!data) return false;
     const hasIdentifiers = !!(
         data.vendor_name && data.vendor_name.trim() &&
         data.invoice_number && data.invoice_number.trim() &&
         data.invoice_date && data.invoice_date.trim()
     );
     // Accept either invoice_value OR taxable_amount (insurance docs use "Net Premium" etc.)
     const hasAmount = (parseFloat(data.invoice_value) || 0) > 0 ||
                       (parseFloat(data.taxable_amount) || 0) > 0;
     return hasIdentifiers && hasAmount;
 }

 // Merge page-2+ data into the base extraction — prefer non-empty/non-zero values; concatenate line items
 mergeExtractions(base, supplement) {
     if (!supplement) return base;
     const merged = { ...base };
     for (const [key, val] of Object.entries(supplement)) {
         if (key === 'line_items') {
             const baseItems = Array.isArray(base.line_items) ? base.line_items : [];
             const suppItems = Array.isArray(val) ? val : [];
             merged.line_items = [...baseItems, ...suppItems];
         } else if (typeof val === 'number') {
             if ((!merged[key] || merged[key] === 0) && val !== 0) merged[key] = val;
         } else if (typeof val === 'string') {
             if ((!merged[key] || merged[key].trim() === '') && val && val.trim()) merged[key] = val;
         }
     }
     return merged;
 }

 // Remove financial summary rows that Claude sometimes includes as line items
 filterSummaryLineItems(lineItems) {
     if (!Array.isArray(lineItems)) return [];
     // Matches summary/tax rows whether they appear at start of description or prefixed with "Output", "Input", etc.
     const summaryPattern = /^(sub.?total|grand total|net total|gross total|final total|amount payable|total payable|payable amount|balance due|balance payable|total amount|total tax|total gst|igst|cgst|sgst|gst\s*@|tax amount|round.?off|total\s+(premium|od|act|basic|tax|gst|amount|payable|due|charges?)|net\s+premium|gross\s+premium|final\s+premium|(output|input)\s+(cgst|sgst|igst|gst))/i;
     return lineItems.filter(item => {
         const desc = (item.description || '').trim();
         if (!desc) return false;
         // Never filter negative-amount items that are not summary labels — they are real discounts/returns
         if ((item.line_total < 0 || item.unit_rate < 0) && desc.length > 2 && !summaryPattern.test(desc)) return true;
         if (summaryPattern.test(desc)) {
             console.log(`   🧹 Filtered summary row from line items: "${desc}"`);
             return false;
         }
         return true;
     });
 }

 getExtractionPrompt() {
     return `You are an expert invoice data extraction system for Indian GST invoices.

IMPORTANT — MULTIPLE INVOICES DETECTION:
If this document contains MORE THAN ONE separate invoice (identified by different invoice numbers, different dates, or a new invoice header starting mid-document), return a JSON ARRAY where each element is one complete invoice object.
If the document contains exactly ONE invoice (even across multiple pages), return a single JSON object (not an array).
Examples:
- 5-page document, pages 1-2 = Invoice A, pages 3-5 = Invoice B → return [{...invoiceA}, {...invoiceB}]
- 7-page document, all pages = one invoice → return {...invoice}

LANGUAGE TRANSLATION RULE:
If any text on the invoice is in a non-English language (Hindi, Marathi, Gujarati, Tamil, Telugu, Kannada, Bengali, or any other language), translate ALL extracted text fields into English before putting them in the JSON. This applies to: vendor_name, customer_name, vendor_address, description, and all line item descriptions. Numeric fields, GST numbers, invoice numbers, and dates must be extracted as-is (no translation needed). Examples:
- "सर्वसाधारण कर (९०% घटक भुक्तासहित)" → "General Tax (90% component inclusive)"
- "जल लाभ कर" → "Water Benefit Tax"
- "स्वच्छाई कर / मलिन. सारण कर" → "Cleanliness Tax / Sewage Tax"
- "नवी मुंबई महानगरपालिका" → "Navi Mumbai Municipal Corporation"

Extract ALL data and return ONLY valid JSON (no markdown, no explanation).

Required structure for each invoice object:
{

Required JSON structure:
{
  "vendor_name": "Full legal name of the supplier/seller",
  "vendor_gstn": "Supplier GSTIN (15-character alphanumeric, uppercase) — may appear near vendor name, at bottom of invoice near PAN number, or in a signature/authorization block",
  "vendor_address": "Complete supplier address as printed",
  "customer_name": "The BUYER / Bill-To party name. If the invoice has both a 'Consignee' (ship-to) and a 'Buyer (if other than consignee)' section, use the BUYER section — NOT the consignee. The buyer is the one who pays; the consignee is just where goods are delivered.",
  "customer_gstn": "GSTIN of the BUYER (Bill-To party), not the consignee. If invoice has separate Buyer and Consignee sections, use Buyer's GSTIN.",
  "invoice_number": "Invoice/bill number exactly as printed",
  "invoice_date": "Invoice date in DD/MM/YYYY format",
  "place_of_supply": "Place of supply state name",
  "taxable_amount": <number: the PRE-TAX subtotal — look for "Taxable Amount", "Taxable Value", "Assessable Value", "Total Taxable" in the GST summary table. For insurance: "Net Premium" before GST. NEVER use "Balance Due", "Amount Payable", "Grand Total", or "Total" as taxable_amount — those are post-tax figures>,
  "igst_amount": <number: IGST tax amount from GST summary, 0 if not present>,
  "igst_rate": <number: IGST rate %, 0 if not present>,
  "cgst_amount": <number: CGST tax amount — look for "CGST", "CGST(9%)", "CGST @" anywhere in the document including premium schedule tables. Sum across all rows if multiple. 0 if not present>,
  "cgst_rate": <number: CGST rate %, 0 if not present>,
  "sgst_amount": <number: SGST/UTGST amount — look for "SGST", "SGST(9%)", "SGST @" anywhere in the document including premium schedule tables. Sum across all rows if multiple. 0 if not present>,
  "sgst_rate": <number: SGST rate %, 0 if not present>,
  "cess_amount": <number: CESS amount, 0 if not present>,
  "cess_rate": <number: CESS rate %, 0 if not present>,
  "tds_amount": <number: TDS deducted, 0 if not present>,
  "tds_rate": <number: TDS rate %, 0 if not present>,
  "round_off": <number: rounding adjustment (can be negative), 0 if not present>,
  "invoice_value": <number: the FINAL payable amount AFTER all taxes — look for "Balance Due", "Grand Total", "Total Amount Payable", "Invoice Total", "Net Payable". For insurance policies: "Gross Premium Paid", "Gross Premium", "Final Premium", "Total Premium Payable". This must equal taxable_amount + all taxes>,
  "description": "Primary product or service description",
  "hsn_sac_code": "HSN or SAC code (primary item)",
  "line_items": [
    {
      "description": "Item description",
      "hsn_sac": "HSN/SAC code",
      "quantity": <number>,
      "unit": "Nos/Kg/Ltr/etc",
      "unit_rate": <number: rate per unit>,
      "line_total": <number: quantity × rate, before tax>,
      "tax_rate": <number: GST % for this item>,
      "tax_amount": <number: total tax on this item>,
      "igst_amount": <number: 0 if intra-state>,
      "cgst_amount": <number: 0 if inter-state>,
      "sgst_amount": <number: 0 if inter-state>
    }
  ]
}

Rules:
- All numeric fields MUST be plain numbers (no ₹, no commas, no currency symbols)
- Indian number format: amounts like "11,66,667" or "1,05,000" use Indian lakh/crore comma placement — convert to plain integers (11,66,667 → 1166667; 1,05,000 → 105000)
- Use 0 for missing numbers; "" for missing strings
- IGST = inter-state; CGST+SGST = intra-state
- Extract every line item visible in the table
- Do NOT fabricate data — only extract what is clearly printed
- CRITICAL: taxable_amount is ALWAYS the pre-tax base (from GST Summary table). invoice_value is ALWAYS the final total after tax (Balance Due / Grand Total). They must NOT be the same number unless tax is 0.
- LINE ITEMS SOURCE RULE: By default, extract ALL individual product/service rows from the main line items table as line_items. EXCEPTION: if the invoice has invoice-level discounts or scheme adjustments that are deducted after the product rows (meaning individual row amounts are pre-discount and therefore inaccurate), AND the invoice also has a "GST Receipt Summary" / "GST Summary" / "HSN Summary" table showing post-discount taxable values per HSN — then use the HSN summary rows as line_items instead (each HSN row = one line item with the correct netted taxable value). How to tell which case applies: if every product row has quantity × rate = amount with no separate discount row affecting totals → use individual rows. If there are discount rows (e.g. "MOP Discount", "Scheme Discount") reducing the subtotal, or if the sum of individual row amounts does NOT equal the taxable_amount in the GST summary → use HSN summary rows. When using HSN summary rows: { description: <short category label for all products under that HSN — never a single product name, never the raw HSN number>, hsn_sac: HSN code, quantity: 1, unit_rate: taxable value, line_total: taxable value, tax_rate: SGST%+CGST% or IGST%, tax_amount: total tax for that HSN }.
- For insurance policy schedules: the pre-tax premium subtotal row (before any GST rows) = taxable_amount. CGST and SGST rows appear below it in the same table — extract those amounts. The final total premium including GST = invoice_value.
- CGST and SGST labels can appear anywhere — inside premium tables, schedule tables, or summary sections. Always scan the ENTIRE document for these labels.
- For service invoices (software, consulting, maintenance): the single line item amount IS the taxable_amount. "Add: CGST @" and "Add: SGST @" rows below it are the tax amounts. The final "Total" is invoice_value.
- Invoice number may be labeled "Invoice No:", "Invoice No", "Bill No", "Reference No" — scan the full header area.
- CONTINUATION PAGES: If the line items table starts with a serial number greater than 1 (e.g., the first visible item is numbered 23, 15, etc.), this is a continuation page of a multi-page invoice. Extract all line items visible on this page normally. The invoice header (number, date, vendor, customer) is still printed at the top of each page — extract it as usual. Do NOT skip line items just because they are a continuation.
- For government/municipal/utility bills (property tax, water charges, electricity, local body levies, etc.) that contain NO mention of GST/CGST/SGST/IGST: the percentage columns are levy/charge rates, NOT GST rates. Set igst_amount=0, cgst_amount=0, sgst_amount=0, and set tax_rate=0 for all line items. The total payable amount = both taxable_amount and invoice_value.
- ONLY populate cgst_amount, sgst_amount, igst_amount when you explicitly see the words "CGST", "SGST", "IGST", or "GST" with a rupee amount on the document.
- BUYER vs CONSIGNEE: Many Indian GST invoices have a "Consignee" (delivery address) AND a separate "Buyer (if other than consignee)" or "Bill To" section. Always use the BUYER/BILL-TO party as customer_name and customer_gstn. The consignee is only the delivery recipient and should be ignored for customer fields.
- Line items may include negative amounts (discounts, returns, adjustments, promotions) — include them as-is with negative values. Any row labelled as a discount, deduction, rebate, scheme, or return that has a negative or bracketed amount must have negative unit_rate and line_total (e.g. -3000, not 0).
- Return ONLY the JSON (object or array), no markdown, no explanation`;
 }

 getSupplementaryPrompt(missingFields) {
     return `This is a continuation page of a multi-page invoice. Extract ONLY the fields listed below if visible on this page.
If any text is in a non-English language (Hindi, Marathi, Gujarati, etc.), translate all text fields to English before returning them.

Missing fields to find: ${missingFields.join(', ')}

IMPORTANT for line_items: Only add items that are actual products/services/or charge components with their own individual descriptions. Do NOT add summary or tax rows (e.g. subtotals, totals, GST amounts, round-offs, or any row that repeats the invoice's aggregate figures) as line items — those are financial summaries that belong in the amount fields above.

Return ONLY a JSON object with this structure (use 0 for missing numbers, "" for missing strings):
{
  "vendor_name": "",
  "vendor_gstn": "",
  "vendor_address": "",
  "customer_name": "",
  "customer_gstn": "",
  "invoice_number": "",
  "invoice_date": "",
  "place_of_supply": "",
  "taxable_amount": <number: base amount before tax — the pre-GST subtotal>,
  "igst_amount": <number: IGST tax amount>,
  "igst_rate": <number: IGST rate %>,
  "cgst_amount": <number: CGST amount>,
  "cgst_rate": <number: CGST rate %>,
  "sgst_amount": <number: SGST amount>,
  "sgst_rate": <number: SGST rate %>,
  "cess_amount": 0,
  "cess_rate": 0,
  "tds_amount": 0,
  "tds_rate": 0,
  "round_off": 0,
  "invoice_value": <number: final amount payable including all taxes — "Grand Total", "Balance Due", "Total Payable", or equivalent>,
  "description": "",
  "hsn_sac_code": "",
  "line_items": []
}

Rules:
- All numeric fields MUST be plain numbers (no ₹, no commas)
- Return ONLY the JSON object, no markdown, no explanation`;
 }

 parseClaudeJSON(rawText) {
     const jsonStr = rawText.replace(/^```(?:json)?\s*/i, '').replace(/\s*```$/, '');
     try {
         return JSON.parse(jsonStr);
     } catch {
         // Try array first, then single object
         const arrMatch = rawText.match(/\[[\s\S]*\]/);
         if (arrMatch) {
             try { return JSON.parse(arrMatch[0]); } catch {}
         }
         const match = rawText.match(/\{[\s\S]*\}/);
         if (!match) throw new Error('Claude response is not valid JSON');
         return JSON.parse(match[0]);
     }
 }

 // Send one image buffer (base64) to Claude and return raw extracted JSON
 async callClaudeWithImage(base64Data, mediaType, prompt) {
     const response = await this.client.messages.create({
         model: this.config.model,
         max_tokens: 4096,
         messages: [{
             role: 'user',
             content: [
                 { type: 'image', source: { type: 'base64', media_type: mediaType, data: base64Data } },
                 { type: 'text', text: prompt }
             ]
         }]
     });
     return response.content[0].text.trim();
 }

 // Convert a PDF page buffer to JPEG base64 (smaller = fewer tokens)
 async pageBufferToBase64(pageBuffer) {
     if (imageOptimizationAvailable) {
         const jpeg = await sharp(pageBuffer).jpeg({ quality: 85 }).toBuffer();
         return { base64: jpeg.toString('base64'), mediaType: 'image/jpeg' };
     }
     return { base64: pageBuffer.toString('base64'), mediaType: 'image/png' };
 }

 // 📄 Process a PDF by converting each page to an image (token-efficient)
 // Only advances to the next page if required fields are still missing
 async processPdfAsImages(filePath, voucherType) {
     const doc = await pdfToImg(filePath, { scale: 1.5 });
     const pageCount = doc.length;
     console.log(`📄 PDF has ${pageCount} page(s) — processing as images`);

     let extractedRaw = null;
     let rawText = '';
     let pageNum = 0;

     for await (const pageBuffer of doc) {
         pageNum++;
         const { base64, mediaType } = await this.pageBufferToBase64(pageBuffer);

         let prompt;
         if (pageNum === 1) {
             prompt = this.getExtractionPrompt();
         } else {
             const missing = [];
             if (!extractedRaw.vendor_name) missing.push('vendor name, GSTIN, address');
             if (!extractedRaw.invoice_number) missing.push('invoice number');
             if (!extractedRaw.invoice_date) missing.push('invoice date');
             if (!(parseFloat(extractedRaw.invoice_value) > 0) && !(parseFloat(extractedRaw.taxable_amount) > 0)) missing.push('total amount (invoice_value / taxable_amount / Net Premium / Final Premium), tax breakdown');
             if (!Array.isArray(extractedRaw.line_items) || extractedRaw.line_items.length === 0) missing.push('line items');
             prompt = this.getSupplementaryPrompt(missing);
         }

         console.log(`   🔍 Page ${pageNum}/${pageCount}${pageNum > 1 ? ' (supplementary)' : ''}...`);
         const rawResponse = await this.callClaudeWithImage(base64, mediaType, prompt);
         const pageData = this.parseClaudeJSON(rawResponse);
         if (Array.isArray(pageData.line_items)) {
             pageData.line_items = this.filterSummaryLineItems(pageData.line_items);
         }

         extractedRaw = pageNum === 1 ? pageData : this.mergeExtractions(extractedRaw, pageData);
         rawText += (pageNum > 1 ? `\n[Page ${pageNum}]\n` : '') + rawResponse;

         if (this.isExtractionComplete(extractedRaw)) {
             if (pageNum < pageCount) {
                 console.log(`✅ All required fields found on page ${pageNum} — skipping ${pageCount - pageNum} remaining page(s)`);
             }
             break;
         }
         if (pageNum < pageCount) console.log(`⚠️  Required fields incomplete after page ${pageNum}, reading page ${pageNum + 1}...`);
     }

     return { extractedRaw, rawText };
 }

 // 🖼️ Process an image file or PDF-as-document directly (used when pdf-to-img is unavailable)
 async processDirectFile(filePath, voucherType) {
     const fileBuffer = fs.readFileSync(filePath);
     const mediaType = this.getMediaType(filePath);

     let base64Data = fileBuffer.toString('base64');
     if (filePath.match(/\.(bmp|tiff)$/i) && imageOptimizationAvailable) {
         const jpeg = await sharp(fileBuffer).jpeg().toBuffer();
         base64Data = jpeg.toString('base64');
     }

     const contentBlock = mediaType === 'application/pdf'
         ? { type: 'document', source: { type: 'base64', media_type: 'application/pdf', data: base64Data } }
         : { type: 'image', source: { type: 'base64', media_type: mediaType, data: base64Data } };

     const response = await this.client.messages.create({
         model: this.config.model,
         max_tokens: 4096,
         messages: [{
             role: 'user',
             content: [contentBlock, { type: 'text', text: this.getExtractionPrompt() }]
         }]
     });

     const rawText = response.content[0].text.trim();
     const parsed = this.parseClaudeJSON(rawText);

     // Claude returned multiple invoices from a single PDF
     if (Array.isArray(parsed)) {
         const invoices = parsed.map(inv => {
             if (Array.isArray(inv.line_items)) inv.line_items = this.filterSummaryLineItems(inv.line_items);
             return inv;
         });
         return { extractedRaw: invoices[0], multipleInvoices: invoices, rawText };
     }

     // Single invoice (normal case)
     if (Array.isArray(parsed.line_items)) {
         parsed.line_items = this.filterSummaryLineItems(parsed.line_items);
     }
     return { extractedRaw: parsed, rawText };
 }

 async processInvoice(filePath, originalName, voucherType = 'Purchase') {
     const startTime = Date.now();
     const isPdf = path.extname(filePath).toLowerCase() === '.pdf';

     try {
         console.log('🤖 Processing invoice with Claude:', originalName);

         // Route: PDF → Claude native document reader (bypasses font rendering issues)
         //        Image → page-by-page image conversion if available, else direct
         let extractedRaw, rawText, multipleInvoices;
         if (isPdf) {
             // Claude reads PDFs natively as documents — more reliable than image conversion
             // which fails for PDFs with non-standard embedded fonts (e.g. Aptos, subsetted Helvetica)
             const directResult = await this.processDirectFile(filePath, voucherType);
             extractedRaw = directResult.extractedRaw;
             rawText = directResult.rawText;
             multipleInvoices = directResult.multipleInvoices; // set if Claude found >1 invoice

             // If native reading gave incomplete results, retry with page-by-page image conversion
             if (!multipleInvoices && !this.isExtractionComplete(extractedRaw) && pdfConversionAvailable) {
                 console.log('⚠️ Direct PDF reading incomplete — retrying with page-by-page image conversion...');
                 const imageResult = await this.processPdfAsImages(filePath, voucherType);
                 extractedRaw = imageResult.extractedRaw;
                 rawText = imageResult.rawText;
             }
         } else {
             // Non-PDF images go directly to Claude vision
             const result = await this.processDirectFile(filePath, voucherType);
             extractedRaw = result.extractedRaw;
             rawText = result.rawText;
             multipleInvoices = result.multipleInvoices;
         }

         const processingTime = Date.now() - startTime;

         // ── Multi-invoice PDF: return all invoices as an array ────────────────────────────
         if (multipleInvoices && multipleInvoices.length > 1) {
             console.log(`📄 Multi-invoice PDF detected: ${multipleInvoices.length} invoices in ${originalName}`);
             const allExtracted = multipleInvoices.map(raw => {
                 const ed = this.convertToInternalFormat(raw, voucherType);
                 const warnings = this.validateGSTINs(ed);
                 if (warnings.length > 0) ed.validation_warnings = JSON.stringify(warnings);
                 return { ...ed, processing_time_ms: processingTime, confidence_score: 0.92, processing_source: 'claude' };
             });
             return { success: true, extractedData: allExtracted[0], multipleInvoices: allExtracted, rawText, processingTime };
         }

         const extractedData = this.convertToInternalFormat(extractedRaw, voucherType);

         const gstinWarnings = this.validateGSTINs(extractedData);
         if (gstinWarnings.length > 0) {
             console.log('⚠️ GSTIN validation warnings:', gstinWarnings);
             extractedData.validation_warnings = JSON.stringify(gstinWarnings);
         }

         console.log('✅ Claude processing complete:', {
             vendor: extractedData.vendor_name,
             amount: extractedData.invoice_value,
             lineItems: extractedData.line_items?.length || 0,
             voucherType: extractedData.voucher_type,
             method: isPdf && pdfConversionAvailable ? 'pdf-as-images' : 'direct'
         });

         return {
             success: true,
             extractedData: {
                 ...extractedData,
                 processing_time_ms: processingTime,
                 confidence_score: 0.92,
                 processing_source: 'claude'
             },
             rawText,
             processingTime
         };

     } catch (error) {
         console.error('❌ Claude processing failed:', error);
         throw error;
     }
 }

 convertToInternalFormat(data, voucherType) {
     const parseNum = (v) => {
         if (v === null || v === undefined || v === '') return 0;
         const n = parseFloat(String(v).replace(/[^0-9.-]/g, ''));
         return isNaN(n) ? 0 : Math.round(n * 100) / 100;
     };

     const lineItems = Array.isArray(data.line_items) ? data.line_items.map(item => {
         const lineTotal = parseNum(item.line_total);
         const taxRate  = parseNum(item.tax_rate);
         let taxAmount  = parseNum(item.tax_amount);

         // Only auto-compute tax if line_total is positive (not a discount row) and tax is genuinely missing
         if (taxAmount === 0 && taxRate > 0 && lineTotal > 0) {
             taxAmount = Math.round((lineTotal * taxRate / 100) * 100) / 100;
         }

         return {
             description: item.description || '',
             hsn_code: item.hsn_sac || item.hsn_code || '',
             quantity: parseNum(item.quantity) || 1,
             unit: item.unit || 'Nos',
             unit_rate: parseNum(item.unit_rate),
             line_total: lineTotal,
             tax_rate: taxRate,
             tax_amount: taxAmount,
             igst_amount: parseNum(item.igst_amount),
             cgst_amount: parseNum(item.cgst_amount),
             sgst_amount: parseNum(item.sgst_amount)
         };
     }) : [];

     const igst = parseNum(data.igst_amount);
     const cgst = parseNum(data.cgst_amount);
     const sgst = parseNum(data.sgst_amount);
     const cess = parseNum(data.cess_amount);

     return {
         invoice_date: data.invoice_date || '',
         invoice_number: data.invoice_number || '',
         vendor_name: data.vendor_name || '',
         vendor_gstn: (data.vendor_gstn || '').trim().toUpperCase(),
         vendor_address: data.vendor_address || '',
         customer_name: data.customer_name || '',
         customer_gstn: (data.customer_gstn || '').trim().toUpperCase(),
         place_of_supply: data.place_of_supply || '',
         taxable_amount: parseNum(data.taxable_amount),
         igst_amount: igst, igst_rate: parseNum(data.igst_rate),
         cgst_amount: cgst, cgst_rate: parseNum(data.cgst_rate),
         sgst_amount: sgst, sgst_rate: parseNum(data.sgst_rate),
         cess_amount: cess, cess_rate: parseNum(data.cess_rate),
         tds_amount: parseNum(data.tds_amount), tds_rate: parseNum(data.tds_rate),
         round_off: parseNum(data.round_off),
         invoice_value: parseNum(data.invoice_value),
         description: data.description || '',
         hsn_sac_code: data.hsn_sac_code || '',
         line_items: lineItems,
         line_items_count: lineItems.length,
         has_line_items: lineItems.length > 0 ? 1 : 0,
         voucher_type: voucherType,
         ledger_name: '',
         total_tax_amount: igst + cgst + sgst + cess,
         category: 'General',
         quantity: lineItems.length > 0 ? lineItems[0].quantity : 1,
         unit_rate: lineItems.length > 0 ? lineItems[0].unit_rate : 0,
         line_item_amount: lineItems.reduce((s, i) => s + i.line_total, 0)
     };
 }

 // Replace visually similar non-ASCII characters (e.g. Cyrillic О → O) in GSTINs
 sanitizeGSTIN(gstin) {
     if (!gstin) return '';
     // Normalize to ASCII: replace common Unicode lookalikes for letters and digits
     return gstin.trim()
         .toUpperCase()
         .replace(/[^\x00-\x7F]/g, (ch) => {
             const map = {
                 'А':'A','В':'B','С':'C','Е':'E','Н':'H','І':'I','Ј':'J',
                 'К':'K','М':'M','О':'O','Р':'P','Q':'Q','Ѕ':'S','Т':'T',
                 'Ʋ':'V','Х':'X','Ү':'Y','Z':'Z','0':'0'
             };
             return map[ch] || ch;
         });
 }

 validateGSTINs(data) {
     const warnings = [];

     if (data.vendor_gstn) {
         const sanitized = this.sanitizeGSTIN(data.vendor_gstn);
         if (sanitized !== data.vendor_gstn) {
             console.log(`🔧 Vendor GSTIN sanitized: "${data.vendor_gstn}" → "${sanitized}"`);
             data.vendor_gstn = sanitized;
         }
         const result = GSTINValidator.validate(data.vendor_gstn);
         if (!result.valid) {
             warnings.push(`Vendor GSTIN "${data.vendor_gstn}" invalid: ${result.error}`);
         } else if (data.vendor_address) {
             const addrState = GSTINValidator.inferStateFromAddress(data.vendor_address);
             if (addrState && addrState !== result.stateName) {
                 warnings.push(`Vendor GSTIN state (${result.stateName}) may not match address state (${addrState}) — please verify`);
             }
         }
     }

     if (data.customer_gstn) {
         const sanitized = this.sanitizeGSTIN(data.customer_gstn);
         if (sanitized !== data.customer_gstn) {
             data.customer_gstn = sanitized;
         }
         const result = GSTINValidator.validate(data.customer_gstn);
         if (!result.valid) warnings.push(`Customer GSTIN "${data.customer_gstn}" invalid: ${result.error}`);
     }
     return warnings;
 }
}

// Enhanced Amount Processor with Parseur data preservation
class EnhancedAmountProcessor {
 static AMOUNT_TOLERANCE = 0.5;
 static VALID_GST_RATES = [0, 3, 5, 12, 18, 28];

 static processAmountsWithEnhancedLogic(extractedData, processingSource = 'unknown') {
     const result = {
         originalData: { ...extractedData },
         processedData: { ...extractedData },
         confidence: 1.0,
         errors: [],
         warnings: [],
         processingSource: processingSource
     };
     
     if (processingSource === 'parseur' && extractedData.confidence_score > 0.9) {
         console.log('🔒 Preserving high-confidence Parseur data, skipping heavy processing');
         result.confidence = extractedData.confidence_score;
         result.warnings.push('Parseur data preserved - skipped recalculation');
         
         this.validateParseurData(result);
         return result;
     }
     
     const complexity = this.detectInvoiceComplexity(result.processedData);
     
     switch (complexity.method) {
         case 'rate_based_correction':
             this.applyRateBasedCorrection(result);
             break;
         case 'line_item_extraction':
             this.applyLineItemExtraction(result);
             break;
         default:
             this.applyRateBasedCorrection(result);
     }
     
     this.forceMathematicalCalculation(result);
     
     if (result.processedData.line_items && result.processedData.line_items.length > 0) {
         this.validateLineItems(result);
     }
     
     this.validateAllCalculations(result);
     
     return result;
 }

 static validateParseurData(result) {
     const data = result.processedData;
     
     if (!data.vendor_name || data.vendor_name === 'Unknown Vendor') {
         result.warnings.push('Vendor name may need review');
     }
     
     if (!data.vendor_gstn || data.vendor_gstn.length !== 15) {
         result.warnings.push('Vendor GSTN may need review');
     }
     
     if (!data.invoice_number) {
         result.warnings.push('Invoice number may need review');
     }
     
     if (data.invoice_value <= 0) {
         result.warnings.push('Invoice value may need review');
     }
     
     if (data.line_items && data.line_items.length > 0) {
         data.line_items.forEach((item, index) => {
             if (!item.description || item.description.trim().length < 3) {
                 result.warnings.push(`Line item ${index + 1} description may need review`);
             }
             if (item.line_total <= 0) {
                 result.warnings.push(`Line item ${index + 1} amount may need review`);
             }
         });
     }
     
     console.log('✅ Parseur data validation complete (no modifications made)');
 }

 static detectInvoiceComplexity(data) {
     const rates = [
         data.cgst_rate || 0,
         data.sgst_rate || 0, 
         data.igst_rate || 0,
         data.cess_rate || 0
     ].filter(rate => rate > 0);
     
     const uniqueRates = [...new Set(rates)];
     
     if (uniqueRates.length <= 1) {
         return {
             type: 'single_rate',
             method: 'rate_based_correction'
         };
     }
     
     if (uniqueRates.length > 1) {
         const hasLineItems = data.line_items && data.line_items.length > 0;
         
         if (hasLineItems) {
             return {
                 type: 'multi_rate_with_line_items',
                 method: 'line_item_extraction'
             };
         } else {
             return {
                 type: 'multi_rate_no_line_items',
                 method: 'tax_summary_extraction'
             };
         }
     }
     
     return {
         type: 'complex',
         method: 'manual_review_required'
     };
 }

 static applyRateBasedCorrection(result) {
     const data = result.processedData;
     const taxableValue = data.taxable_amount || data.taxable_value || 0;
     
     if (taxableValue <= 0) {
         result.errors.push('No taxable value found');
         return;
     }
     
     if (data.cgst_rate > 0) {
         const expectedCGST = this.roundToTwoDecimals((taxableValue * data.cgst_rate) / 100);
         const actualCGST = data.cgst_amount || 0;
         
         if (Math.abs(expectedCGST - actualCGST) > this.AMOUNT_TOLERANCE) {
             result.processedData.cgst_amount = expectedCGST;
         }
     }
     
     if (data.sgst_rate > 0) {
         const expectedSGST = this.roundToTwoDecimals((taxableValue * data.sgst_rate) / 100);
         const actualSGST = data.sgst_amount || 0;
         
         if (Math.abs(expectedSGST - actualSGST) > this.AMOUNT_TOLERANCE) {
             result.processedData.sgst_amount = expectedSGST;
         }
     }
     
     if (data.igst_rate > 0) {
         const expectedIGST = this.roundToTwoDecimals((taxableValue * data.igst_rate) / 100);
         const actualIGST = data.igst_amount || 0;
         
         if (Math.abs(expectedIGST - actualIGST) > this.AMOUNT_TOLERANCE) {
             result.processedData.igst_amount = expectedIGST;
         }
     }
     
     if (data.cess_rate > 0) {
         const expectedCESS = this.roundToTwoDecimals((taxableValue * data.cess_rate) / 100);
         const actualCESS = data.cess_amount || 0;
         
         if (Math.abs(expectedCESS - actualCESS) > this.AMOUNT_TOLERANCE) {
             result.processedData.cess_amount = expectedCESS;
         }
     }
     
     this.validateCGSTSGSTParity(result);
 }

 static applyLineItemExtraction(result) {
     const data = result.processedData;
     
     if (!data.line_items || data.line_items.length === 0) {
         this.applyTaxSummaryExtraction(result);
         return;
     }
     
     let totalCGST = 0, totalSGST = 0, totalIGST = 0, totalCESS = 0;
     let totalTaxable = 0;
     
     try {
         for (const lineItem of data.line_items) {
             const lineItemTaxable = lineItem.line_total || 0;
             const lineItemCGSTRate = lineItem.cgst_rate || 0;
             const lineItemSGSTRate = lineItem.sgst_rate || 0;
             const lineItemIGSTRate = lineItem.igst_rate || lineItem.tax_rate || 0;
             const lineItemCESSRate = lineItem.cess_rate || 0;
             
             if (lineItemTaxable > 0) {
                 totalTaxable += lineItemTaxable;
                 totalCGST += (lineItemTaxable * lineItemCGSTRate) / 100;
                 totalSGST += (lineItemTaxable * lineItemSGSTRate) / 100;
                 totalIGST += (lineItemTaxable * lineItemIGSTRate) / 100;
                 totalCESS += (lineItemTaxable * lineItemCESSRate) / 100;
             }
         }
         
         if (totalTaxable > 0) {
             result.processedData.taxable_amount = this.roundToTwoDecimals(totalTaxable);
         }
         result.processedData.cgst_amount = this.roundToTwoDecimals(totalCGST);
         result.processedData.sgst_amount = this.roundToTwoDecimals(totalSGST);
         result.processedData.igst_amount = this.roundToTwoDecimals(totalIGST);
         result.processedData.cess_amount = this.roundToTwoDecimals(totalCESS);
         
     } catch (error) {
         this.applyTaxSummaryExtraction(result);
     }
 }

 static validateLineItems(result) {
     const lineItems = result.processedData.line_items;
     let lineItemIssues = [];
     
     lineItems.forEach((item, index) => {
         if (item.quantity > 0 && item.unit_rate > 0) {
             const expectedTotal = this.roundToTwoDecimals(item.quantity * item.unit_rate);
             if (Math.abs(item.line_total - expectedTotal) > this.AMOUNT_TOLERANCE) {
                 lineItemIssues.push(`Line ${index + 1}: Amount calculation mismatch`);
                 if (result.processingSource !== 'parseur') {
                     item.line_total = expectedTotal;
                 }
             }
         }
         
         if (item.line_total > 0 && item.tax_rate > 0) {
             const expectedTax = this.roundToTwoDecimals((item.line_total * item.tax_rate) / 100);
             if (Math.abs(item.tax_amount - expectedTax) > this.AMOUNT_TOLERANCE) {
                 lineItemIssues.push(`Line ${index + 1}: Tax calculation mismatch`);
                 if (result.processingSource !== 'parseur') {
                     item.tax_amount = expectedTax;
                 }
             }
         }
         
         if (!item.description || item.description.trim().length < 3) {
             lineItemIssues.push(`Line ${index + 1}: Missing or invalid description`);
         }
         
         if (item.line_total <= 0) {
             lineItemIssues.push(`Line ${index + 1}: Invalid line total`);
         }
     });
     
     if (lineItemIssues.length > 0) {
         result.warnings = result.warnings.concat(lineItemIssues);
     }
 }

 static applyTaxSummaryExtraction(result) {
     // Keep extracted values with validation
 }

 static forceMathematicalCalculation(result) {
     if (result.processingSource === 'parseur') {
         return;
     }
     
     const data = result.processedData;
     
     const cgst = data.cgst_amount || 0;
     const sgst = data.sgst_amount || 0;
     const igst = data.igst_amount || 0;
     const cess = data.cess_amount || 0;
     
     const calculatedTotalTax = this.roundToTwoDecimals(cgst + sgst + igst + cess);
     
     if (data.total_tax_amount !== calculatedTotalTax) {
         result.processedData.total_tax_amount = calculatedTotalTax;
     }
     
     const taxableValue = data.taxable_amount || data.taxable_value || 0;
     const totalTax = data.total_tax_amount || 0;
     const roundOff = data.round_off || 0;
     const tdsAmount = data.tds_amount || 0;
     
     const calculatedInvoiceValue = this.roundToTwoDecimals(taxableValue + totalTax + roundOff - tdsAmount);
     
     if (Math.abs((data.invoice_value || 0) - calculatedInvoiceValue) > this.AMOUNT_TOLERANCE) {
         result.processedData.invoice_value = calculatedInvoiceValue;
         result.processedData.amount = calculatedInvoiceValue;
     }
 }

 static validateCGSTSGSTParity(result) {
     if (result.processingSource === 'parseur') {
         return;
     }
     
     const cgstAmount = result.processedData.cgst_amount || 0;
     const sgstAmount = result.processedData.sgst_amount || 0;
     
     if (cgstAmount > 0 && sgstAmount > 0) {
         const difference = Math.abs(cgstAmount - sgstAmount);
         
         if (difference > this.AMOUNT_TOLERANCE) {
             const averageAmount = this.roundToTwoDecimals((cgstAmount + sgstAmount) / 2);
             result.processedData.cgst_amount = averageAmount;
             result.processedData.sgst_amount = averageAmount;
             result.warnings.push('CGST and SGST amounts corrected for parity');
         }
     }
 }

 static validateAllCalculations(result) {
     const data = result.processedData;
     
     const tolerance = result.processingSource === 'parseur' ? 1.0 : this.AMOUNT_TOLERANCE;
     
     const expectedInvoiceValue = this.roundToTwoDecimals(
         (data.taxable_amount || data.taxable_value || 0) + 
         (data.total_tax_amount || 0) + 
         (data.round_off || 0) - 
         (data.tds_amount || 0)
     );
     
     if (Math.abs((data.invoice_value || 0) - expectedInvoiceValue) > tolerance) {
         result.errors.push(`Invoice value validation failed`);
         result.confidence -= 0.2;
     }
     
     const expectedTotalTax = this.roundToTwoDecimals(
         (data.cgst_amount || 0) + (data.sgst_amount || 0) + (data.igst_amount || 0) + (data.cess_amount || 0)
     );
     
     if (Math.abs((data.total_tax_amount || 0) - expectedTotalTax) > tolerance) {
         result.errors.push(`Total tax validation failed`);
         result.confidence -= 0.15;
     }
     
     const rates = [data.cgst_rate, data.sgst_rate, data.igst_rate].filter(rate => rate > 0);
     rates.forEach(rate => {
         if (!this.VALID_GST_RATES.includes(rate)) {
             result.warnings.push(`Invalid GST rate: ${rate}%`);
             result.confidence -= 0.1;
         }
     });
     
     if (data.line_items && data.line_items.length > 0) {
         const lineItemsTotal = data.line_items.reduce((sum, item) => sum + (item.line_total || 0), 0);
         const invoiceTaxable = data.taxable_amount || 0;
         
         if (Math.abs(lineItemsTotal - invoiceTaxable) > tolerance) {
             result.warnings.push(`Line items total (₹${lineItemsTotal}) doesn't match invoice taxable amount (₹${invoiceTaxable})`);
             result.confidence -= 0.1;
         }
     }
 }

 static roundToTwoDecimals(value) {
     return Math.round((value + Number.EPSILON) * 100) / 100;
 }

 static parseAmount(amountStr) {
     if (!amountStr) return 0;
     
     let cleaned = amountStr.toString()
         .replace(/[₹,Rs\s]/g, '')
         .replace(/[^\d\.]/g, '');
     
     const decimalParts = cleaned.split('.');
     if (decimalParts.length > 2) {
         cleaned = decimalParts.slice(0, -1).join('') + '.' + decimalParts[decimalParts.length - 1];
     }
     
     const amount = parseFloat(cleaned) || 0;
     return this.roundToTwoDecimals(amount);
 }
}

// Unified Document Processor with Voucher Type Support
class UnifiedDocumentProcessor {
 static async processInvoice(filePath, originalName, voucherType = 'Purchase') {
     const startTime = Date.now();

     // 1️⃣ PRIMARY: Claude AI
     if (CLAUDE_CONFIG.enabled) {
         try {
             console.log('🤖 Attempting Claude AI processing...');
             const claudeResult = await claudeProcessor.processInvoice(filePath, originalName, voucherType);

             if (claudeResult.success) {
                 const finalData = this.applyEnhancedGSTLogic(claudeResult.extractedData, 'claude');
                 // Pass multipleInvoices through if Claude found more than one invoice in this PDF
                 const multipleInvoices = claudeResult.multipleInvoices
                     ? claudeResult.multipleInvoices.map(inv => this.applyEnhancedGSTLogic(inv, 'claude'))
                     : undefined;
                 return {
                     success: true,
                     extractedData: finalData,
                     multipleInvoices,
                     rawText: claudeResult.rawText,
                     processingSource: 'claude',
                     processingTime: claudeResult.processingTime
                 };
             }
         } catch (claudeError) {
             console.log('⚠️ Claude processing failed, falling back to Document AI:', claudeError.message);
         }
     }

     // 2️⃣ FIRST FALLBACK: Parseur (if still configured)
     if (PARSEUR_CONFIG.enabled) {
         try {
             console.log('🔄 Falling back to Parseur processing...');
             const parseurResult = await parseurProcessor.processInvoiceWithParseur(filePath, originalName, voucherType);

             if (parseurResult.success) {
                 const finalData = this.applyEnhancedGSTLogic(parseurResult.extractedData, 'parseur');
                 return {
                     success: true,
                     extractedData: finalData,
                     rawText: parseurResult.rawText,
                     processingSource: 'parseur',
                     processingTime: parseurResult.processingTime
                 };
             }
         } catch (parseurError) {
             console.log('⚠️ Parseur fallback failed, trying Document AI:', parseurError.message);
         }
     }

     // 3️⃣ SECOND FALLBACK: Document AI
     try {
         console.log('🔄 Processing with Document AI...');
         const documentAIResult = await this.processWithDocumentAI(filePath, originalName, voucherType);

         if (documentAIResult.success) {
             const finalData = this.applyEnhancedGSTLogic(documentAIResult.extractedData, 'document_ai');
             return {
                 success: true,
                 extractedData: finalData,
                 rawText: documentAIResult.rawText,
                 processingSource: 'document_ai',
                 processingTime: Date.now() - startTime
             };
         }
     } catch (docAIError) {
         console.log('⚠️ Document AI failed:', docAIError.message);
     }

     throw new Error('All processing methods failed (Claude, Parseur, Document AI)');
 }

 static async processWithDocumentAI(filePath, originalName, voucherType = 'Purchase') {
     console.log('⚠️ Document AI processing not fully implemented in this example');
     
     return {
         success: true,
         extractedData: {
             vendor_name: 'Unknown Vendor',
             invoice_number: 'Unknown',
             invoice_value: 0,
             voucher_type: voucherType,
             line_items: [],
             confidence_score: 0.5
         },
         rawText: 'Document AI processing placeholder'
     };
 }

 static applyEnhancedGSTLogic(extractedData, processingSource) {
     return extractedData;
 }
}

// 🎯 FIXED: Enhanced Helper function for Tally conversion with proper voucher structure - NO STANDARDIZATION HERE
function convertToTallyFormat(rows) {
 const tallyData = [];
 
 // 🗓️ Helper function to format dates properly for Tally
 const formatDateForTally = (dateStr) => {
     if (!dateStr) return new Date().toLocaleDateString('en-GB');
     const date = new Date(dateStr);
     return date.toLocaleDateString('en-GB', {
         day: '2-digit',
         month: '2-digit', 
         year: 'numeric'
     }); // Ensures DD/MM/YYYY format that Tally loves
 };
     
 rows.forEach((row, invoiceIndex) => {
     const invoiceDate = formatDateForTally(row.invoice_date);
     const voucherType = row.voucher_type || 'Purchase';
     const invoiceNumber = row.invoice_number || `INV-${invoiceIndex + 1}`;
     const vendorName = row.vendor_name || 'Unknown Vendor';
     const vendorAddress = row.vendor_address || '';
     const voucherNarration = `${voucherType} from ${vendorName}`;
     
     // 🆕 SPECIAL HANDLING FOR JOURNAL VOUCHERS
     if (voucherType === 'Journal') {
         try {
             // Parse the journal entries from extracted_text
             const journalData = JSON.parse(row.extracted_text || '{}');
             
             if (journalData.debit_entries && journalData.credit_entries) {
                 // Add debit entries
                 journalData.debit_entries.forEach((debitEntry, index) => {
                     tallyData.push({
                         "Voucher Date": invoiceDate,
                         "Voucher Type Name": "Journal",
                         "Voucher Number": invoiceNumber,
                         "Buyer/Supplier - Address": '',
                         "Buyer/Supplier - Pincode": '',
                         "Ledger Name": debitEntry.ledger_name,
                         "Ledger Amount": debitEntry.amount,
                         "Ledger Amount Dr/Cr": 'Dr',
                         "Item Name": '',
                         "Billed Quantity": '',
                         "Item Rate": '',
                         "Item Rate per": '',
                         "Item Amount": '',
                         "Change Mode ": '',
                         "Voucher Narration": journalData.narration || voucherNarration
                     });
                 });
                 
                 // Add credit entries
                 journalData.credit_entries.forEach((creditEntry, index) => {
                     tallyData.push({
                         "Voucher Date": invoiceDate,
                         "Voucher Type Name": "Journal",
                         "Voucher Number": invoiceNumber,
                         "Buyer/Supplier - Address": '',
                         "Buyer/Supplier - Pincode": '',
                         "Ledger Name": creditEntry.ledger_name,
                         "Ledger Amount": creditEntry.amount,
                         "Ledger Amount Dr/Cr": 'Cr',
                         "Item Name": '',
                         "Billed Quantity": '',
                         "Item Rate": '',
                         "Item Rate per": '',
                         "Item Amount": '',
                         "Change Mode ": '',
                         "Voucher Narration": journalData.narration || voucherNarration
                     });
                 });
                 
                 return; // Skip the regular processing for journal vouchers
             }
         } catch (error) {
             console.log('Failed to parse journal data for:', invoiceNumber);
             // Fall through to regular processing
         }
     }
     
     // REGULAR PURCHASE/SALES VOUCHER PROCESSING
     // Parse line items safely
     let lineItems = [];
     try {
         if (row.line_items && row.line_items !== '[]' && row.line_items !== 'null') {
             const parsed = JSON.parse(row.line_items);
             lineItems = Array.isArray(parsed) ? parsed : [];
         }
     } catch (e) {
         console.log('Failed to parse line items for invoice:', invoiceNumber);
         lineItems = [];
     }
     
     // Drop line items with no amount (e.g. insurance coverage descriptions)
     const meaningfulLineItems = lineItems.filter(item => (item.line_total || 0) > 0 || (item.unit_rate || 0) > 0);

     // If no meaningful line items, fall back to a single summary entry using the invoice total
     if (meaningfulLineItems.length === 0) {
         lineItems = [{
             description: row.description || 'Standard Purchase Item',
             quantity: row.quantity || 1,
             unit_rate: row.unit_rate || row.taxable_amount || row.invoice_value || 0,
             line_total: row.taxable_amount || row.invoice_value || 0,
             tax_rate: row.igst_rate || row.cgst_rate && row.cgst_rate * 2 || 18,
             tax_amount: row.total_tax_amount || 0
         }];
     } else {
         lineItems = meaningfulLineItems;
     }
     
     const drEntries = [];
     
     // 1️⃣ Create entries for each line item (Individual Items)
     lineItems.forEach((item, itemIndex) => {
         const itemName = item.description || `Item ${itemIndex + 1}`;
         const taxRate = item.tax_rate || 18;
         
         drEntries.push({
             "Voucher Date": invoiceDate,
             "Voucher Type Name": voucherType,
             "Voucher Number": invoiceNumber,
             "Buyer/Supplier - Address": vendorAddress,
             "Buyer/Supplier - Pincode": '',
             "Ledger Name": voucherType === 'Purchase' ? 'Purchase' : 'Sales',
             "Ledger Amount": item.line_total || 0,
             "Ledger Amount Dr/Cr": voucherType === 'Purchase' ? 'Dr' : 'Cr',
             "Item Name": itemName,
             "Billed Quantity": item.quantity || 1,
             "Item Rate": item.unit_rate || 0,
             "Item Rate per": 'Nos',
             "Item Amount": item.line_total || 0,
             "Change Mode ": '',
             "Voucher Narration": voucherNarration
         });
     });
     
     // 2️⃣ Add tax entries (CGST) - 🎯 GENERATE DYNAMIC NAMES WITHOUT STANDARDIZATION
     if (row.cgst_amount > 0) {
         drEntries.push({
             "Voucher Date": invoiceDate,
             "Voucher Type Name": voucherType,
             "Voucher Number": invoiceNumber,
             "Buyer/Supplier - Address": '',
             "Buyer/Supplier - Pincode": '',
             "Ledger Name": voucherType === 'Purchase' ? `Input CGST @ ${row.cgst_rate || 9}%` : `Output CGST @ ${row.cgst_rate || 9}%`,
             "Ledger Amount": row.cgst_amount,
             "Ledger Amount Dr/Cr": voucherType === 'Purchase' ? 'Dr' : 'Cr',
             "Item Name": '',
             "Billed Quantity": '',
             "Item Rate": '',
             "Item Rate per": '',
             "Item Amount": '',
             "Change Mode ": '',
             "Voucher Narration": voucherNarration
         });
     }
     
     // Add tax entries (SGST) - 🎯 GENERATE DYNAMIC NAMES WITHOUT STANDARDIZATION
     if (row.sgst_amount > 0) {
         drEntries.push({
             "Voucher Date": invoiceDate,
             "Voucher Type Name": voucherType,
             "Voucher Number": invoiceNumber,
             "Buyer/Supplier - Address": '',
             "Buyer/Supplier - Pincode": '',
             "Ledger Name": voucherType === 'Purchase' ? `Input SGST @ ${row.sgst_rate || 9}%` : `Output SGST @ ${row.sgst_rate || 9}%`,
             "Ledger Amount": row.sgst_amount,
             "Ledger Amount Dr/Cr": voucherType === 'Purchase' ? 'Dr' : 'Cr',
             "Item Name": '',
             "Billed Quantity": '',
             "Item Rate": '',
             "Item Rate per": '',
             "Item Amount": '',
             "Change Mode ": '',
             "Voucher Narration": voucherNarration
         });
     }
     
     // Add tax entries (IGST) - 🎯 GENERATE DYNAMIC NAMES WITHOUT STANDARDIZATION
     if (row.igst_amount > 0) {
         drEntries.push({
             "Voucher Date": invoiceDate,
             "Voucher Type Name": voucherType,
             "Voucher Number": invoiceNumber,
             "Buyer/Supplier - Address": '',
             "Buyer/Supplier - Pincode": '',
             "Ledger Name": voucherType === 'Purchase' ? `Input IGST @ ${row.igst_rate || 18}%` : `Output IGST @ ${row.igst_rate || 18}%`,
             "Ledger Amount": row.igst_amount,
             "Ledger Amount Dr/Cr": voucherType === 'Purchase' ? 'Dr' : 'Cr',
             "Item Name": '',
             "Billed Quantity": '',
             "Item Rate": '',
             "Item Rate per": '',
             "Item Amount": '',
             "Change Mode ": '',
             "Voucher Narration": voucherNarration
         });
     }
     
     // 3️⃣ Calculate rounding off
     const invoiceValue = row.invoice_value || 0;
     const totalDrAmount = drEntries.reduce((sum, entry) => sum + (entry["Ledger Amount"] || 0), 0);
     const roundingDiff = invoiceValue - totalDrAmount;
     
     // Add rounding off entry if there's a significant difference
     if (Math.abs(roundingDiff) > 0.01) {
         drEntries.push({
             "Voucher Date": invoiceDate,
             "Voucher Type Name": voucherType,
             "Voucher Number": invoiceNumber,
             "Buyer/Supplier - Address": '',
             "Buyer/Supplier - Pincode": '',
             "Ledger Name": "Round Off",
             "Ledger Amount": Math.abs(roundingDiff),
             "Ledger Amount Dr/Cr": roundingDiff > 0 ? (voucherType === 'Purchase' ? "Dr" : "Cr") : (voucherType === 'Purchase' ? "Cr" : "Dr"),
             "Item Name": '',
             "Billed Quantity": '',
             "Item Rate": '',
             "Item Rate per": '',
             "Item Amount": '',
             "Change Mode ": '',
             "Voucher Narration": voucherNarration
         });
     }
     
     // 4️⃣ Add credit/debit entry for vendor/customer
     const crEntry = {
         "Voucher Date": invoiceDate,
         "Voucher Type Name": voucherType,
         "Voucher Number": invoiceNumber,
         "Buyer/Supplier - Address": vendorAddress,
         "Buyer/Supplier - Pincode": '',
         "Ledger Name": vendorName,
         "Ledger Amount": invoiceValue,
         "Ledger Amount Dr/Cr": voucherType === 'Purchase' ? 'Cr' : 'Dr',
         "Item Name": '',
         "Billed Quantity": '',
         "Item Rate": '',
         "Item Rate per": '',
         "Item Amount": '',
         "Change Mode ": '',
         "Voucher Narration": voucherNarration
     };
     
     // Add all entries for this invoice to the final output
     tallyData.push(...drEntries, crEntry);
 });
 
 return tallyData;
}

// Initialize Claude processor (primary)
let claudeProcessor;
if (CLAUDE_CONFIG.enabled) {
 claudeProcessor = new ClaudeInvoiceProcessor(CLAUDE_CONFIG);
 console.log('✅ Claude AI processor initialized (primary extractor)');
} else {
 console.log('⚠️  Claude AI disabled — set ANTHROPIC_API_KEY in .env to enable');
}

// Initialize Parseur processor (first fallback)
let parseurProcessor;
if (PARSEUR_CONFIG.enabled) {
 parseurProcessor = new ParseurInvoiceProcessor(PARSEUR_CONFIG);
 console.log('✅ Parseur processor initialized (first fallback)');
}

// Basic endpoints
app.get('/', (req, res) => {
 res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

app.get('/api/expenses', (req, res) => {
 const sql = `SELECT *,
     CASE WHEN (file_path IS NOT NULL OR parseur_document_id IS NOT NULL) THEN 1 ELSE 0 END as can_reprocess
     FROM expenses ORDER BY created_at DESC`;
 
 db.all(sql, (err, rows) => {
     if (err) {
         console.error('Query failed:', err);
         return res.status(500).json({ error: 'Failed to fetch expenses' });
     }
     
     // Parse line_items JSON for each row and fix counts
     const processedRows = rows.map(row => {
         try {
             // Parse line_items if it exists and is a string
             if (row.line_items && typeof row.line_items === 'string') {
                 row.line_items = JSON.parse(row.line_items);
             } else if (!row.line_items) {
                 row.line_items = [];
             }
             
             // Make sure it's an array
             if (!Array.isArray(row.line_items)) {
                 row.line_items = [];
             }
             
             // Fix the counts based on actual data
             row.line_items_count = row.line_items.length;
             row.has_line_items = row.line_items.length > 0 ? 1 : 0;
             
         } catch (e) {
             console.warn(`Failed to parse line_items for expense ${row.id}:`, e);
             row.line_items = [];
             row.line_items_count = 0;
             row.has_line_items = 0;
         }
         return row;
     });
     
     res.json(processedRows);
 });
});

// 🧠 NEW REPROCESSING ENDPOINT WITH PARSEUR DOCUMENT ID
app.post('/api/expenses/:id/reprocess', async (req, res) => {
 try {
     const expenseId = req.params.id;

     const existingExpense = await new Promise((resolve, reject) => {
         const sql = 'SELECT * FROM expenses WHERE id = ?';
         db.get(sql, [expenseId], (err, row) => {
             if (err) reject(err);
             else resolve(row);
         });
     });

     if (!existingExpense) {
         return res.status(404).json({ success: false, error: 'Expense not found' });
     }

     const canUseFile = existingExpense.file_path && fs.existsSync(existingExpense.file_path);
     const canUseParseur = !!(existingExpense.parseur_document_id && PARSEUR_CONFIG.enabled && parseurProcessor);

     if (!canUseFile && !canUseParseur) {
         return res.status(400).json({
             success: false,
             error: 'Reprocessing not available',
             details: 'Original file no longer exists and no Parseur document ID is stored'
         });
     }

     let reprocessResult;
     let usedSource;

     // Prefer Claude (file-based) if the file is available
     if (canUseFile && CLAUDE_CONFIG.enabled && claudeProcessor) {
         console.log(`🤖 Reprocessing expense ${expenseId} with Claude (file: ${existingExpense.file_path})`);
         reprocessResult = await claudeProcessor.processInvoice(
             existingExpense.file_path,
             existingExpense.original_filename || path.basename(existingExpense.file_path),
             existingExpense.voucher_type || 'Purchase'
         );
         usedSource = 'claude_reprocessed';
     } else if (canUseParseur) {
         console.log(`🔄 Reprocessing expense ${expenseId} with Parseur document ID: ${existingExpense.parseur_document_id}`);
         reprocessResult = await parseurProcessor.reprocessDocument(existingExpense.parseur_document_id);
         usedSource = 'parseur_reprocessed';
     } else if (canUseFile) {
         // File exists but Claude not configured — run through full pipeline
         reprocessResult = await UnifiedDocumentProcessor.processInvoice(
             existingExpense.file_path,
             existingExpense.original_filename || path.basename(existingExpense.file_path),
             existingExpense.voucher_type || 'Purchase'
         );
         usedSource = reprocessResult.processingSource + '_reprocessed';
     }

     if (!reprocessResult || !reprocessResult.success) {
         throw new Error(`Reprocessing failed: ${reprocessResult?.error || 'unknown error'}`);
     }

     const extractedData = reprocessResult.extractedData;
     
     // Preserve original voucher type and file information
     extractedData.voucher_type = existingExpense.voucher_type;
     extractedData.entry_type = 'reprocessed';
     
     console.log(`📊 Reprocessing result:`, {
         vendor: extractedData.vendor_name,
         amount: extractedData.invoice_value,
         lineItems: extractedData.line_items?.length || 0,
         confidence: Math.round((extractedData.confidence_score || 0) * 100) + '%',
         voucherType: extractedData.voucher_type
     });
     
     // Update database with reprocessed data
     const updateSql = `UPDATE expenses SET 
         invoice_date = ?, invoice_number = ?, vendor_name = ?, vendor_gstn = ?, 
         customer_name = ?, customer_gstn = ?, place_of_supply = ?,
         taxable_amount = ?, igst_amount = ?, cgst_amount = ?, sgst_amount = ?, cess_amount = ?, round_off = ?, invoice_value = ?,
         tds_rate = ?, tds_amount = ?, description = ?, hsn_sac_code = ?, line_item_amount = ?, quantity = ?, unit_rate = ?,
         ledger_name = ?, vendor_address = ?, total_tax_amount = ?,
         cgst_rate = ?, sgst_rate = ?, igst_rate = ?, cess_rate = ?, category = ?,
         extracted_text = ?, confidence_score = ?, amount_confidence = ?, 
         processing_time_ms = ?, status = ?, processing_source = ?,
         line_items = ?, line_items_count = ?, has_line_items = ?, table_structure_confidence = ?,
         processing_method = ?, validation_errors = ?, validation_warnings = ?, entry_type = ?
         WHERE id = ?`;
     
     const values = [
         prepareDateValue(extractedData.invoice_date, 'invoice_date'),
         prepareValue(extractedData.invoice_number, 'invoice_number'),
         prepareValue(extractedData.vendor_name, 'vendor_name') || 'Unknown Vendor',
         prepareValue(extractedData.vendor_gstn, 'vendor_gstn'),
         prepareValue(extractedData.customer_name, 'customer_name'),
         prepareValue(extractedData.customer_gstn, 'customer_gstn'),
         prepareValue(extractedData.place_of_supply, 'place_of_supply'),
         prepareNumericValue(extractedData.taxable_amount || extractedData.taxable_value, 'taxable_amount'),
         prepareNumericValue(extractedData.igst_amount, 'igst_amount'),
         prepareNumericValue(extractedData.cgst_amount, 'cgst_amount'),
         prepareNumericValue(extractedData.sgst_amount, 'sgst_amount'),
         prepareNumericValue(extractedData.cess_amount, 'cess_amount'),
         prepareNumericValue(extractedData.round_off, 'round_off'),
         prepareNumericValue(extractedData.invoice_value || extractedData.amount, 'invoice_value'),
         prepareNumericValue(extractedData.tds_rate, 'tds_rate'),
         prepareNumericValue(extractedData.tds_amount, 'tds_amount'),
         prepareValue(extractedData.description, 'description'),
         prepareValue(extractedData.hsn_sac_code, 'hsn_sac_code'),
         prepareNumericValue(extractedData.line_item_amount, 'line_item_amount'),
         prepareNumericValue(extractedData.quantity, 'quantity'),
         prepareNumericValue(extractedData.unit_rate, 'unit_rate'),
         prepareValue(extractedData.ledger_name, 'ledger_name'),
         prepareValue(extractedData.vendor_address, 'vendor_address'),
         prepareNumericValue(extractedData.total_tax_amount, 'total_tax_amount'),
         prepareNumericValue(extractedData.cgst_rate, 'cgst_rate'),
         prepareNumericValue(extractedData.sgst_rate, 'sgst_rate'),
         prepareNumericValue(extractedData.igst_rate, 'igst_rate'),
         prepareNumericValue(extractedData.cess_rate, 'cess_rate'),
         prepareValue(extractedData.category, 'category') || 'General',
         prepareValue(reprocessResult.rawText, 'extracted_text'),
         prepareNumericValue(extractedData.confidence_score, 'confidence_score', 0.95),
         prepareNumericValue(extractedData.amount_confidence, 'amount_confidence', 0.95),
         prepareNumericValue(extractedData.processing_time_ms, 'processing_time_ms', 500),
         'pending_review',
         usedSource,
         JSON.stringify(extractedData.line_items || []),
         prepareNumericValue(extractedData.line_items_count, 'line_items_count', 0),
         prepareBooleanValue(extractedData.has_line_items, 'has_line_items'),
         prepareNumericValue(extractedData.table_structure_confidence, 'table_structure_confidence', 0),
         'reprocess',
         prepareValue(extractedData.validation_errors, 'validation_errors') || '[]',
         prepareValue(extractedData.validation_warnings, 'validation_warnings') || '[]',
         'reprocessed',
         expenseId
     ];

     await new Promise((resolve, reject) => {
         db.run(updateSql, values, function(err) {
             if (err) reject(err);
             else resolve(this.changes);
         });
     });

     console.log(`✅ Expense ${expenseId} reprocessed successfully via ${usedSource}`);

     res.json({
         success: true,
         message: `Expense reprocessed successfully via ${usedSource}`,
         expense: {
             id: expenseId,
             invoice_number: extractedData.invoice_number,
             vendor_name: extractedData.vendor_name,
             invoice_value: extractedData.invoice_value || extractedData.amount,
             line_items_count: extractedData.line_items_count || 0,
             processing_source: usedSource,
             confidence: Math.round((extractedData.confidence_score || 0) * 100),
             voucher_type: extractedData.voucher_type,
             entry_type: 'reprocessed',
             reprocessed: true,
             processing_time_ms: extractedData.processing_time_ms || 500
         },
         reprocessing: {
             original_processing_source: existingExpense.processing_source,
             reprocessed_with: usedSource
         }
     });
     
 } catch (error) {
     console.error('❌ Expense reprocessing failed:', error);
     res.status(500).json({
         success: false,
         error: 'Expense reprocessing failed',
         details: error.message
     });
 }
});

// 🔐 ENHANCED UPLOAD INVOICE ENDPOINT WITH CONTENT-BASED DUPLICATE DETECTION
app.post('/api/upload-invoice', upload.single('receipt'), async (req, res) => {
 try {
     if (!req.file) {
         return res.status(400).json({ success: false, error: 'No file uploaded' });
     }
     
     const voucherType = req.body.voucher_type || 'Purchase';
     
     console.log('📤 Processing invoice:', req.file.originalname, 'Voucher Type:', voucherType);
     
     // 🔐 STEP 1: ENHANCED DUPLICATE DETECTION WITH CONTENT-BASED HASHING
     console.log('🔍 Performing enhanced duplicate detection...');
     const duplicateCheck = await DuplicateDetector.validateUpload(req.file);
     
     if (!duplicateCheck.isValid) {
         // File was already deleted in validateUpload
         if (duplicateCheck.isDuplicate) {
             return res.status(409).json({
                 success: false,
                 error: 'Duplicate file content detected',
                 message: duplicateCheck.message,
                 details: duplicateCheck.details,
                 duplicateType: duplicateCheck.duplicateType,
                 fileHash: duplicateCheck.fileHash,
                 existingFile: duplicateCheck.existingFile,
                 reprocessAllowed: duplicateCheck.reprocessAllowed,
                 reprocessMessage: duplicateCheck.reprocessMessage,
                 // 🧠 Provide reprocessing instructions if available
                 actions: duplicateCheck.reprocessAllowed ? [
                     {
                         action: 'reprocess',
                         description: 'Reprocess the existing document with updated settings',
                         endpoint: `/api/expenses/${duplicateCheck.existingFile.id}/reprocess`,
                         method: 'POST'
                     }
                 ] : []
             });
         } else {
             return res.status(409).json({
                 success: false,
                 error: 'File validation failed',
                 message: duplicateCheck.message,
                 details: duplicateCheck.details
             });
         }
     }
     
     console.log('✅ File passed enhanced duplicate detection');
     
     // Handle filename conflicts with warning
     if (duplicateCheck.isNameConflict) {
         console.log('⚠️ Filename conflict detected but allowing upload (different content)');
     }
     
     const fileExtension = path.extname(req.file.originalname).toLowerCase();
     const documentType = fileExtension === '.pdf' ? 'pdf' : 'image';
     
     // Process with enhanced system including voucher type
     const result = await UnifiedDocumentProcessor.processInvoice(req.file.path, req.file.originalname, voucherType);

     // ── Multi-invoice PDF: insert all invoices and return combined response ──────────────
     if (result.multipleInvoices && result.multipleInvoices.length > 1) {
         console.log(`📄 Single-upload multi-invoice PDF: inserting ${result.multipleInvoices.length} invoices`);
         const insertedInvoices = [];
         const skippedDuplicates = [];

         for (const invRaw of result.multipleInvoices) {
             invRaw.voucher_type = voucherType;
             invRaw.entry_type = 'upload';
             if (invRaw.line_items && Array.isArray(invRaw.line_items)) {
                 invRaw.line_items_count = invRaw.line_items.length;
                 invRaw.has_line_items = invRaw.line_items_count > 0;
             }

             // Duplicate check per invoice
             const dup = await new Promise((resolve, reject) => {
                 if (!invRaw.invoice_number) return resolve(null);
                 db.get(`SELECT id, invoice_number, vendor_name FROM expenses WHERE invoice_number = ? AND invoice_number != ''`,
                     [invRaw.invoice_number], (err, row) => err ? reject(err) : resolve(row));
             });
             if (dup) {
                 console.log(`⚠️  Invoice ${invRaw.invoice_number} already exists (ID: ${dup.id}) — skipping`);
                 skippedDuplicates.push({ invoice_number: invRaw.invoice_number, vendor_name: invRaw.vendor_name });
                 continue;
             }

             const insertSql = `INSERT INTO expenses (
                 invoice_date, invoice_number, vendor_name, vendor_gstn, customer_name, customer_gstn, place_of_supply,
                 taxable_amount, igst_amount, cgst_amount, sgst_amount, cess_amount, round_off, invoice_value,
                 tds_rate, tds_amount, description, hsn_sac_code, line_item_amount, quantity, unit_rate,
                 voucher_type, ledger_name, vendor_address, total_tax_amount,
                 cgst_rate, sgst_rate, igst_rate, cess_rate, category, file_path, original_filename, file_hash,
                 parseur_document_id, extracted_text,
                 confidence_score, amount_confidence, document_type, processing_time_ms, status, processing_source,
                 line_items, line_items_count, has_line_items, table_structure_confidence,
                 processing_method, validation_errors, validation_warnings, entry_type
             ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`;
             const vals = [
                 prepareDateValue(invRaw.invoice_date, 'invoice_date'),
                 prepareValue(invRaw.invoice_number, 'invoice_number'),
                 prepareValue(invRaw.vendor_name, 'vendor_name') || 'Unknown Vendor',
                 prepareValue(invRaw.vendor_gstn, 'vendor_gstn'),
                 prepareValue(invRaw.customer_name, 'customer_name'),
                 prepareValue(invRaw.customer_gstn, 'customer_gstn'),
                 prepareValue(invRaw.place_of_supply, 'place_of_supply'),
                 prepareNumericValue(invRaw.taxable_amount || invRaw.taxable_value, 'taxable_amount'),
                 prepareNumericValue(invRaw.igst_amount, 'igst_amount'),
                 prepareNumericValue(invRaw.cgst_amount, 'cgst_amount'),
                 prepareNumericValue(invRaw.sgst_amount, 'sgst_amount'),
                 prepareNumericValue(invRaw.cess_amount, 'cess_amount'),
                 prepareNumericValue(invRaw.round_off, 'round_off'),
                 prepareNumericValue(invRaw.invoice_value || invRaw.amount, 'invoice_value'),
                 prepareNumericValue(invRaw.tds_rate, 'tds_rate'),
                 prepareNumericValue(invRaw.tds_amount, 'tds_amount'),
                 prepareValue(invRaw.description, 'description'),
                 prepareValue(invRaw.hsn_sac_code, 'hsn_sac_code'),
                 prepareNumericValue(invRaw.line_item_amount, 'line_item_amount'),
                 prepareNumericValue(invRaw.quantity, 'quantity'),
                 prepareNumericValue(invRaw.unit_rate, 'unit_rate'),
                 prepareValue(invRaw.voucher_type, 'voucher_type') || voucherType,
                 prepareValue(invRaw.ledger_name, 'ledger_name'),
                 prepareValue(invRaw.vendor_address, 'vendor_address'),
                 prepareNumericValue(invRaw.total_tax_amount, 'total_tax_amount'),
                 prepareNumericValue(invRaw.cgst_rate, 'cgst_rate'),
                 prepareNumericValue(invRaw.sgst_rate, 'sgst_rate'),
                 prepareNumericValue(invRaw.igst_rate, 'igst_rate'),
                 prepareNumericValue(invRaw.cess_rate, 'cess_rate'),
                 prepareValue(invRaw.category, 'category') || 'General',
                 req.file.path,
                 req.file.originalname,
                 duplicateCheck.fileHash,
                 prepareValue(invRaw.parseur_document_id, 'parseur_document_id'),
                 prepareValue(result.rawText, 'extracted_text'),
                 prepareNumericValue(invRaw.confidence_score, 'confidence_score', 0.8),
                 prepareNumericValue(invRaw.amount_confidence, 'amount_confidence', 0.8),
                 documentType,
                 prepareNumericValue(invRaw.processing_time_ms, 'processing_time_ms', 0),
                 'pending_review',
                 prepareValue(invRaw.processing_source || result.processingSource, 'processing_source'),
                 JSON.stringify(invRaw.line_items || []),
                 prepareNumericValue(invRaw.line_items_count, 'line_items_count', 0),
                 prepareBooleanValue(invRaw.has_line_items, 'has_line_items'),
                 prepareNumericValue(invRaw.table_structure_confidence, 'table_structure_confidence', 0),
                 prepareValue(invRaw.processing_method, 'processing_method') || 'auto',
                 prepareValue(invRaw.validation_errors, 'validation_errors') || '[]',
                 prepareValue(invRaw.validation_warnings, 'validation_warnings') || '[]',
                 'upload'
             ];
             const newId = await new Promise((resolve, reject) => {
                 db.run(insertSql, vals, function(err) { err ? reject(err) : resolve(this.lastID); });
             });
             console.log(`✅ Invoice ${invRaw.invoice_number} saved, ID: ${newId}`);
             insertedInvoices.push({ id: newId, invoice_number: invRaw.invoice_number, vendor_name: invRaw.vendor_name, invoice_value: invRaw.invoice_value });
         }

         return res.json({
             success: true,
             message: `Multi-invoice PDF: ${insertedInvoices.length} invoices saved${skippedDuplicates.length ? `, ${skippedDuplicates.length} already existed` : ''}`,
             multiple_invoices: true,
             invoices: insertedInvoices,
             duplicates_skipped: skippedDuplicates,
             // Keep backward-compatible `expense` field pointing to first inserted invoice
             expense: insertedInvoices[0] ? { ...insertedInvoices[0], line_items_count: result.multipleInvoices[0]?.line_items_count || 0 } : null
         });
     }

     const extractedData = result.extractedData;

     // Ensure voucher type is set
     extractedData.voucher_type = voucherType;
     extractedData.entry_type = 'upload';

     // 🔧 FIX: Calculate line items count properly
     if (extractedData.line_items && Array.isArray(extractedData.line_items)) {
         extractedData.line_items_count = extractedData.line_items.length;
         extractedData.has_line_items = extractedData.line_items_count > 0;
     } else {
         extractedData.line_items_count = 0;
         extractedData.has_line_items = false;
     }

     console.log(`📊 Processing result (${result.processingSource}):`, {
         vendor: extractedData.vendor_name,
         amount: extractedData.invoice_value,
         lineItems: extractedData.line_items_count,
         confidence: Math.round((extractedData.confidence_score || 0) * 100) + '%',
         source: result.processingSource,
         voucherType: extractedData.voucher_type,
         parseurDocId: extractedData.parseur_document_id || 'N/A'
     });

     // Additional duplicate checking based on invoice content (invoice number only)
     const contentDuplicateCheck = await new Promise((resolve, reject) => {
         if (!extractedData.invoice_number || extractedData.invoice_number === '' || extractedData.invoice_number === 'null') return resolve(null);
         db.get(`SELECT * FROM expenses WHERE invoice_number = ? AND invoice_number != '' AND invoice_number != 'null'`,
             [extractedData.invoice_number], (err, row) => err ? reject(err) : resolve(row));
     });

     if (contentDuplicateCheck) {
         // Clean up uploaded file
         await DuplicateDetector.safeDeleteFile(req.file.path);

         return res.status(409).json({
             success: false,
             error: 'Duplicate invoice detected',
             details: `Invoice ${extractedData.invoice_number} from ${extractedData.vendor_name} already exists`,
             existing: contentDuplicateCheck
         });
     }
     
     // Insert with enhanced schema including parseur_document_id
     const insertSql = `INSERT INTO expenses (
         invoice_date, invoice_number, vendor_name, vendor_gstn, customer_name, customer_gstn, place_of_supply,
         taxable_amount, igst_amount, cgst_amount, sgst_amount, cess_amount, round_off, invoice_value,
         tds_rate, tds_amount, description, hsn_sac_code, line_item_amount, quantity, unit_rate,
         voucher_type, ledger_name, vendor_address, total_tax_amount,
         cgst_rate, sgst_rate, igst_rate, cess_rate, category, file_path, original_filename, file_hash, 
         parseur_document_id, extracted_text,
         confidence_score, amount_confidence, document_type, processing_time_ms, status, processing_source,
         line_items, line_items_count, has_line_items, table_structure_confidence,
         processing_method, validation_errors, validation_warnings, entry_type
     ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`;
     
     const values = [
         prepareDateValue(extractedData.invoice_date, 'invoice_date'),
         prepareValue(extractedData.invoice_number, 'invoice_number'),
         prepareValue(extractedData.vendor_name, 'vendor_name') || 'Unknown Vendor',
         prepareValue(extractedData.vendor_gstn, 'vendor_gstn'),
         prepareValue(extractedData.customer_name, 'customer_name'),
         prepareValue(extractedData.customer_gstn, 'customer_gstn'),
         prepareValue(extractedData.place_of_supply, 'place_of_supply'),
         prepareNumericValue(extractedData.taxable_amount || extractedData.taxable_value, 'taxable_amount'),
         prepareNumericValue(extractedData.igst_amount, 'igst_amount'),
         prepareNumericValue(extractedData.cgst_amount, 'cgst_amount'),
         prepareNumericValue(extractedData.sgst_amount, 'sgst_amount'),
         prepareNumericValue(extractedData.cess_amount, 'cess_amount'),
         prepareNumericValue(extractedData.round_off, 'round_off'),
         prepareNumericValue(extractedData.invoice_value || extractedData.amount, 'invoice_value'),
         prepareNumericValue(extractedData.tds_rate, 'tds_rate'),
         prepareNumericValue(extractedData.tds_amount, 'tds_amount'),
         prepareValue(extractedData.description, 'description'),
         prepareValue(extractedData.hsn_sac_code, 'hsn_sac_code'),
         prepareNumericValue(extractedData.line_item_amount, 'line_item_amount'),
         prepareNumericValue(extractedData.quantity, 'quantity'),
         prepareNumericValue(extractedData.unit_rate, 'unit_rate'),
         prepareValue(extractedData.voucher_type, 'voucher_type') || 'Purchase',
         prepareValue(extractedData.ledger_name, 'ledger_name'),
         prepareValue(extractedData.vendor_address, 'vendor_address'),
         prepareNumericValue(extractedData.total_tax_amount, 'total_tax_amount'),
         prepareNumericValue(extractedData.cgst_rate, 'cgst_rate'),
         prepareNumericValue(extractedData.sgst_rate, 'sgst_rate'),
         prepareNumericValue(extractedData.igst_rate, 'igst_rate'),
         prepareNumericValue(extractedData.cess_rate, 'cess_rate'),
         prepareValue(extractedData.category, 'category') || 'General',
         req.file.path,
         req.file.originalname, // Store original filename
         duplicateCheck.fileHash, // Store computed SHA-256 hash
         prepareValue(extractedData.parseur_document_id, 'parseur_document_id'), // 🧠 Store for reprocessing
         prepareValue(result.rawText, 'extracted_text'),
         prepareNumericValue(extractedData.confidence_score, 'confidence_score', 0.8),
         prepareNumericValue(extractedData.amount_confidence, 'amount_confidence', 0.8),
         documentType,
         prepareNumericValue(extractedData.processing_time_ms, 'processing_time_ms', 0),
         'pending_review',
         prepareValue(extractedData.processing_source || result.processingSource, 'processing_source'),
         JSON.stringify(extractedData.line_items || []),
         prepareNumericValue(extractedData.line_items_count, 'line_items_count', 0),
         prepareBooleanValue(extractedData.has_line_items, 'has_line_items'),
         prepareNumericValue(extractedData.table_structure_confidence, 'table_structure_confidence', 0),
         prepareValue(extractedData.processing_method, 'processing_method') || 'auto',
         prepareValue(extractedData.validation_errors, 'validation_errors') || '[]',
         prepareValue(extractedData.validation_warnings, 'validation_warnings') || '[]',
         prepareValue(extractedData.entry_type, 'entry_type') || 'upload'
     ];
     
     db.run(insertSql, values, function(err) {
         if (err) {
             console.error('❌ Insert failed:', err);
             // Clean up uploaded file on database error
             DuplicateDetector.safeDeleteFile(req.file.path);
             return res.status(500).json({ success: false, error: 'Failed to save expense data', details: err.message });
         }
         
         console.log('✅ Invoice processed successfully with enhanced duplicate detection, ID:', this.lastID);
         
         // Enhanced response with reprocessing info
         res.json({
             success: true,
             message: `Invoice processed successfully`,
             expense: { 
                 id: this.lastID, 
                 invoice_number: extractedData.invoice_number,
                 vendor_name: extractedData.vendor_name,
                 vendor_gstn: extractedData.vendor_gstn,
                 customer_name: extractedData.customer_name,
                 customer_gstn: extractedData.customer_gstn,
                 invoice_value: extractedData.invoice_value || extractedData.amount,
                 total_tax_amount: extractedData.total_tax_amount,
                 invoice_date: extractedData.invoice_date,
                 line_items_count: extractedData.line_items_count || 0,
                 has_line_items: extractedData.has_line_items || false,
                 table_confidence: Math.round((extractedData.table_structure_confidence || 0) * 100),
                 processing_source: result.processingSource,
                 confidence: Math.round((extractedData.confidence_score || 0) * 100),
                 voucher_type: extractedData.voucher_type,
                 entry_type: extractedData.entry_type,
                 // 🧠 Reprocessing information
                 parseur_document_id: extractedData.parseur_document_id || null,
                 can_reprocess: !!(extractedData.parseur_document_id),
                 reprocess_available: !!(extractedData.parseur_document_id),
                 duplicate_check: {
                     content_hash_verified: true,
                     filename_conflict_detected: duplicateCheck.isNameConflict || false,
                     file_hash: duplicateCheck.fileHash.substring(0, 8) + '...',
                     validation_passed: true,
                     duplicate_detection_method: 'SHA-256 content-based'
                 }
             },
             // Include filename conflict warning if applicable
             warnings: duplicateCheck.isNameConflict ? [duplicateCheck.warning] : []
         });
     });
     
 } catch (error) {
     console.error('❌ Invoice processing failed:', error);
     
     // Clean up uploaded file on error
     if (req.file && req.file.path) {
         await DuplicateDetector.safeDeleteFile(req.file.path);
     }
     
     res.status(500).json({
         success: false,
         error: 'Invoice processing failed',
         details: error.message
     });
 }
});

// 🆕 NEW MANUAL ENTRY ENDPOINT
app.post('/api/manual-entry', async (req, res) => {
 try {
     const data = req.body;
     
     console.log('✏️ Processing manual entry:', data);
     
     // Validate required fields
     if (!data.vendor_name || !data.invoice_number || !data.invoice_date) {
         return res.status(400).json({
             success: false,
             error: 'Missing required fields: vendor_name, invoice_number, invoice_date'
         });
     }
     
     // Validate and process line items
     const lineItems = data.line_items || [];
     let lineItemsValidation = { isValid: true, processedItems: [], errors: [] };
     
     if (lineItems.length > 0) {
         lineItemsValidation = LineItemsValidator.validateAndProcessLineItems(lineItems);
         
         if (!lineItemsValidation.isValid) {
             return res.status(400).json({
                 success: false,
                 error: 'Line items validation failed',
                 validationErrors: lineItemsValidation.errors
             });
         }
     }
     
     // Calculate totals from line items
     let calculatedTotals = { totalTaxable: 0, totalTax: 0, grandTotal: 0 };
     if (lineItemsValidation.processedItems.length > 0) {
         calculatedTotals = LineItemsValidator.calculateTotals(lineItemsValidation.processedItems);
         console.log(`📊 Manual entry totals calculated: ₹${calculatedTotals.totalTaxable} + ₹${calculatedTotals.totalTax} = ₹${calculatedTotals.grandTotal}`);
     }
     
     // Use calculated values or manual input
     const taxableAmount = calculatedTotals.totalTaxable > 0 ? calculatedTotals.totalTaxable : (parseFloat(data.taxable_amount) || 0);
     const totalTaxAmount = calculatedTotals.totalTax > 0 ? calculatedTotals.totalTax : (parseFloat(data.total_tax_amount) || 0);
     
     // Calculate individual tax amounts based on rates or use provided values
     let cgstAmount = parseFloat(data.cgst_amount) || 0;
     let sgstAmount = parseFloat(data.sgst_amount) || 0;
     let igstAmount = parseFloat(data.igst_amount) || 0;
     
     // If no individual tax amounts provided but total tax is available, split based on interstate/intrastate
     if (totalTaxAmount > 0 && (cgstAmount + sgstAmount + igstAmount) === 0) {
         const customerGSTN = data.customer_gstn || '';
         const vendorGSTN = data.vendor_gstn || '';
         
         // Determine if interstate (different state codes) or intrastate
         const isInterstate = customerGSTN.length >= 2 && vendorGSTN.length >= 2 && 
                            customerGSTN.substring(0, 2) !== vendorGSTN.substring(0, 2);
         
         if (isInterstate) {
             igstAmount = totalTaxAmount;
         } else {
             cgstAmount = totalTaxAmount / 2;
             sgstAmount = totalTaxAmount / 2;
         }
     }
     
     // Calculate final invoice value
     const roundOff = parseFloat(data.round_off) || 0;
     const tdsAmount = parseFloat(data.tds_amount) || 0;
     const finalInvoiceValue = taxableAmount + totalTaxAmount + roundOff - tdsAmount;
     
     // Check for duplicates
     const duplicateCheck = await new Promise((resolve, reject) => {
         const sql = `SELECT * FROM expenses 
             WHERE invoice_number = ? AND vendor_name = ? AND ABS(invoice_value - ?) < 1`;
         
         db.get(sql, [data.invoice_number, data.vendor_name, finalInvoiceValue], (err, row) => {
             if (err) reject(err);
             else resolve(row);
         });
     });
     
     if (duplicateCheck) {
         return res.status(409).json({
             success: false,
             error: 'Duplicate manual entry detected',
             details: `Invoice ${data.invoice_number} from ${data.vendor_name} already exists`,
             existing: duplicateCheck
         });
     }
     
     // Prepare data for insertion (no parseur_document_id for manual entries)
     const insertSql = `INSERT INTO expenses (
         invoice_date, invoice_number, vendor_name, vendor_gstn, customer_name, customer_gstn, place_of_supply,
         taxable_amount, igst_amount, cgst_amount, sgst_amount, cess_amount, round_off, invoice_value,
         tds_rate, tds_amount, description, hsn_sac_code, line_item_amount, quantity, unit_rate,
         voucher_type, ledger_name, vendor_address, total_tax_amount,
         cgst_rate, sgst_rate, igst_rate, cess_rate, category, file_path, original_filename, file_hash, 
         parseur_document_id, extracted_text,
         confidence_score, amount_confidence, document_type, processing_time_ms, status, processing_source,
         line_items, line_items_count, has_line_items, table_structure_confidence,
         processing_method, validation_errors, validation_warnings, entry_type
     ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`;
     
     const values = [
         prepareDateValue(data.invoice_date, 'invoice_date'),
         prepareValue(data.invoice_number, 'invoice_number'),
         prepareValue(data.vendor_name, 'vendor_name'),
         prepareValue(data.vendor_gstn, 'vendor_gstn'),
         prepareValue(data.customer_name, 'customer_name'),
         prepareValue(data.customer_gstn, 'customer_gstn'),
         prepareValue(data.place_of_supply, 'place_of_supply'),
         prepareNumericValue(taxableAmount, 'taxable_amount'),
         prepareNumericValue(igstAmount, 'igst_amount'),
         prepareNumericValue(cgstAmount, 'cgst_amount'),
         prepareNumericValue(sgstAmount, 'sgst_amount'),
         prepareNumericValue(data.cess_amount, 'cess_amount'),
         prepareNumericValue(roundOff, 'round_off'),
         prepareNumericValue(finalInvoiceValue, 'invoice_value'),
         prepareNumericValue(data.tds_rate, 'tds_rate'),
         prepareNumericValue(tdsAmount, 'tds_amount'),
         prepareValue(data.description, 'description'),
         prepareValue(data.hsn_sac_code, 'hsn_sac_code'),
         prepareNumericValue(data.line_item_amount, 'line_item_amount'),
         prepareNumericValue(data.quantity, 'quantity'),
         prepareNumericValue(data.unit_rate, 'unit_rate'),
         prepareValue(data.voucher_type, 'voucher_type') || 'Purchase',
         prepareValue(data.ledger_name, 'ledger_name'),
         prepareValue(data.vendor_address, 'vendor_address'),
         prepareNumericValue(totalTaxAmount, 'total_tax_amount'),
         prepareNumericValue(data.cgst_rate, 'cgst_rate'),
         prepareNumericValue(data.sgst_rate, 'sgst_rate'),
         prepareNumericValue(data.igst_rate, 'igst_rate'),
         prepareNumericValue(data.cess_rate, 'cess_rate'),
         prepareValue(data.category, 'category') || 'General',
         '', // file_path - empty for manual entry
         '', // original_filename - empty for manual entry
         '', // file_hash - empty for manual entry
         null, // parseur_document_id - null for manual entry (cannot be reprocessed)
         '', // extracted_text - empty for manual entry
         1.0, // confidence_score - 100% for manual entry
         1.0, // amount_confidence - 100% for manual entry
         'manual', // document_type
         0, // processing_time_ms
         'pending_review', // status
         'manual', // processing_source
         JSON.stringify(lineItemsValidation.processedItems),
         lineItemsValidation.processedItems.length,
         lineItemsValidation.processedItems.length > 0,
         lineItemsValidation.processedItems.length > 0 ? 1.0 : 0, // table_structure_confidence
         'manual', // processing_method
         JSON.stringify([]), // validation_errors
         JSON.stringify([]), // validation_warnings
         'manual' // entry_type
     ];
     
     db.run(insertSql, values, function(err) {
         if (err) {
             console.error('❌ Manual entry insert failed:', err);
             return res.status(500).json({ 
                 success: false, 
                 error: 'Failed to save manual entry',
                 details: err.message 
             });
         }
         
         console.log('✅ Manual entry created successfully, ID:', this.lastID);
         
         res.json({
             success: true,
             message: 'Manual entry created successfully',
             expense: {
                 id: this.lastID,
                 invoice_number: data.invoice_number,
                 vendor_name: data.vendor_name,
                 invoice_value: finalInvoiceValue,
                 voucher_type: data.voucher_type || 'Purchase',
                 line_items_count: lineItemsValidation.processedItems.length,
                 has_line_items: lineItemsValidation.processedItems.length > 0,
                 entry_type: 'manual',
                 confidence: 100,
                 can_reprocess: false, // Manual entries cannot be reprocessed
                 parseur_document_id: null
             }
         });
     });
     
 } catch (error) {
     console.error('❌ Manual entry processing failed:', error);
     res.status(500).json({
         success: false,
         error: 'Manual entry processing failed',
         details: error.message
     });
 }
});

// 🆕 NEW PAYMENT JOURNAL ENTRY ENDPOINT FOR PAYMENT TAB
app.post('/api/payment-entry', async (req, res) => {
  try {
      const data = req.body;
      
      console.log('💰 Processing payment journal entry:', data);
      
      // Validate the journal entry
      const validation = JournalVoucherProcessor.validateJournalEntry(data);
      
      if (!validation.isValid) {
          return res.status(400).json({
              success: false,
              error: 'Journal entry validation failed',
              validationErrors: validation.errors
          });
      }
      
      // Process the journal entry
      const processedJournal = JournalVoucherProcessor.processJournalEntry(data);
      
      // Convert to expense format for unified storage
      const expenseFormatData = JournalVoucherProcessor.convertToExpenseFormat(processedJournal);
      
      // Check for duplicates based on journal number and date
      const duplicateCheck = await new Promise((resolve, reject) => {
          const sql = `SELECT * FROM expenses 
              WHERE invoice_number = ? AND voucher_type = 'Journal' AND invoice_date = ?`;
          
          db.get(sql, [expenseFormatData.invoice_number, expenseFormatData.invoice_date], (err, row) => {
              if (err) reject(err);
              else resolve(row);
          });
      });
      
      if (duplicateCheck) {
          return res.status(409).json({
              success: false,
              error: 'Duplicate journal entry detected',
              details: `Journal entry ${expenseFormatData.invoice_number} for ${expenseFormatData.invoice_date} already exists`,
              existing: duplicateCheck
          });
      }
      
      // Insert the journal entry into the expenses table
      const insertSql = `INSERT INTO expenses (
          invoice_date, invoice_number, vendor_name, vendor_gstn, customer_name, customer_gstn, place_of_supply,
          taxable_amount, igst_amount, cgst_amount, sgst_amount, cess_amount, round_off, invoice_value,
          tds_rate, tds_amount, description, hsn_sac_code, line_item_amount, quantity, unit_rate,
          voucher_type, ledger_name, vendor_address, total_tax_amount,
          cgst_rate, sgst_rate, igst_rate, cess_rate, category, file_path, original_filename, file_hash, 
          parseur_document_id, extracted_text,
          confidence_score, amount_confidence, document_type, processing_time_ms, status, processing_source,
          line_items, line_items_count, has_line_items, table_structure_confidence,
          processing_method, validation_errors, validation_warnings, entry_type
      ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`;
      
      const values = [
          prepareDateValue(expenseFormatData.invoice_date, 'invoice_date'),
          prepareValue(expenseFormatData.invoice_number, 'invoice_number'),
          prepareValue(expenseFormatData.vendor_name, 'vendor_name'),
          prepareValue(expenseFormatData.vendor_gstn, 'vendor_gstn'),
          prepareValue(expenseFormatData.customer_name, 'customer_name'),
          prepareValue(expenseFormatData.customer_gstn, 'customer_gstn'),
          prepareValue(expenseFormatData.place_of_supply, 'place_of_supply'),
          prepareNumericValue(expenseFormatData.taxable_amount, 'taxable_amount'),
          prepareNumericValue(expenseFormatData.igst_amount, 'igst_amount'),
          prepareNumericValue(expenseFormatData.cgst_amount, 'cgst_amount'),
          prepareNumericValue(expenseFormatData.sgst_amount, 'sgst_amount'),
          prepareNumericValue(expenseFormatData.cess_amount, 'cess_amount'),
          prepareNumericValue(expenseFormatData.round_off, 'round_off'),
          prepareNumericValue(expenseFormatData.invoice_value, 'invoice_value'),
          prepareNumericValue(expenseFormatData.tds_rate, 'tds_rate'),
          prepareNumericValue(expenseFormatData.tds_amount, 'tds_amount'),
          prepareValue(expenseFormatData.description, 'description'),
          prepareValue(expenseFormatData.hsn_sac_code, 'hsn_sac_code'),
          prepareNumericValue(expenseFormatData.line_item_amount, 'line_item_amount'),
          prepareNumericValue(expenseFormatData.quantity, 'quantity'),
          prepareNumericValue(expenseFormatData.unit_rate, 'unit_rate'),
          prepareValue(expenseFormatData.voucher_type, 'voucher_type'),
          prepareValue(expenseFormatData.ledger_name, 'ledger_name'),
          prepareValue(expenseFormatData.vendor_address, 'vendor_address'),
          prepareNumericValue(expenseFormatData.total_tax_amount, 'total_tax_amount'),
          prepareNumericValue(expenseFormatData.cgst_rate, 'cgst_rate'),
          prepareNumericValue(expenseFormatData.sgst_rate, 'sgst_rate'),
          prepareNumericValue(expenseFormatData.igst_rate, 'igst_rate'),
          prepareNumericValue(expenseFormatData.cess_rate, 'cess_rate'),
          prepareValue(expenseFormatData.category, 'category'),
          prepareValue(expenseFormatData.file_path, 'file_path'),
          prepareValue(expenseFormatData.original_filename, 'original_filename'),
          prepareValue(expenseFormatData.file_hash, 'file_hash'),
          expenseFormatData.parseur_document_id, // null for journal entries
          prepareValue(expenseFormatData.extracted_text, 'extracted_text'),
          prepareNumericValue(expenseFormatData.confidence_score, 'confidence_score'),
          prepareNumericValue(expenseFormatData.amount_confidence, 'amount_confidence'),
          prepareValue(expenseFormatData.document_type, 'document_type'),
          prepareNumericValue(expenseFormatData.processing_time_ms, 'processing_time_ms'),
          prepareValue(expenseFormatData.status, 'status'),
          prepareValue(expenseFormatData.processing_source, 'processing_source'),
          prepareValue(expenseFormatData.line_items, 'line_items'),
          prepareNumericValue(expenseFormatData.line_items_count, 'line_items_count'),
          prepareBooleanValue(expenseFormatData.has_line_items, 'has_line_items'),
          prepareNumericValue(expenseFormatData.table_structure_confidence, 'table_structure_confidence'),
          prepareValue(expenseFormatData.processing_method, 'processing_method'),
          prepareValue(expenseFormatData.validation_errors, 'validation_errors'),
          prepareValue(expenseFormatData.validation_warnings, 'validation_warnings'),
          prepareValue(expenseFormatData.entry_type, 'entry_type')
      ];
      
      db.run(insertSql, values, function(err) {
          if (err) {
              console.error('❌ Payment journal entry insert failed:', err);
              return res.status(500).json({ 
                  success: false, 
                  error: 'Failed to save payment journal entry',
                  details: err.message 
              });
          }
          
          console.log('✅ Payment journal entry created successfully, ID:', this.lastID);
          
          res.json({
              success: true,
              message: 'Payment journal entry created successfully',
              expense: {
                  id: this.lastID,
                  journal_number: expenseFormatData.invoice_number,
                  voucher_date: expenseFormatData.invoice_date,
                  total_amount: expenseFormatData.invoice_value,
                  voucher_type: 'Journal',
                  debit_entries: processedJournal.debit_entries,
                  credit_entries: processedJournal.credit_entries,
                  narration: processedJournal.narration,
                  line_items_count: expenseFormatData.line_items_count,
                  has_line_items: expenseFormatData.has_line_items,
                  entry_type: 'payment_journal',
                  confidence: 100,
                  can_reprocess: false, // Journal entries cannot be reprocessed
                  parseur_document_id: null,
                  status: 'pending_review'
              },
              journal_details: {
                  debit_total: processedJournal.debit_entries.reduce((sum, entry) => sum + entry.amount, 0),
                  credit_total: processedJournal.credit_entries.reduce((sum, entry) => sum + entry.amount, 0),
                  entries_count: processedJournal.debit_entries.length + processedJournal.credit_entries.length,
                  balanced: true
              },
              validation: {
                  warnings: validation.warnings
              }
          });
      });
      
  } catch (error) {
      console.error('❌ Payment journal entry processing failed:', error);
      res.status(500).json({
          success: false,
          error: 'Payment journal entry processing failed',
          details: error.message
      });
  }
});

// Enhanced line items update endpoint with comprehensive validation
app.put('/api/expenses/:id/line-items', (req, res) => {
  const expenseId = req.params.id;
  const { line_items } = req.body;
  
  console.log(`🔧 Updating line items for expense ${expenseId}:`, line_items?.length || 0, 'items');
  
  // Validate and process line items
  const validation = LineItemsValidator.validateAndProcessLineItems(line_items);
  
  if (!validation.isValid) {
      return res.status(400).json({ 
          success: false,
          error: 'Line items validation failed', 
          validationErrors: validation.errors 
      });
  }
  
  // Calculate totals from processed line items
  const totals = LineItemsValidator.calculateTotals(validation.processedItems);
  
  // Update database with processed line items and calculated totals
  // invoice_value is recalculated from line items so the Review card stays in sync
  const sql = `UPDATE expenses SET
      line_items = ?,
      line_items_count = ?,
      has_line_items = ?,
      taxable_amount = ?,
      total_tax_amount = ?,
      invoice_value = ?,
      table_structure_confidence = ?,
      status = 'pending_review'
      WHERE id = ?`;

  const values = [
      JSON.stringify(validation.processedItems),
      validation.processedItems.length,
      validation.processedItems.length > 0,
      totals.totalTaxable,
      totals.totalTax,
      totals.grandTotal,
      validation.processedItems.length > 0 ? 0.98 : 0, // High confidence for manually edited items
      expenseId
  ];
  
  db.run(sql, values, function(err) {
      if (err) {
          console.error('❌ Line items update failed:', err);
          return res.status(500).json({ 
              success: false, 
              error: 'Failed to update line items',
              details: err.message 
          });
      }
      
      if (this.changes === 0) {
          return res.status(404).json({ 
              success: false, 
              error: 'Invoice not found' 
          });
      }
      
      console.log(`✅ Line items updated successfully for expense ${expenseId}:`, {
          items: validation.processedItems.length,
          totalTaxable: totals.totalTaxable,
          totalTax: totals.totalTax
      });
      
      res.json({
          success: true,
          message: 'Line items updated successfully',
          line_items_count: validation.processedItems.length,
          totals_updated: {
              taxable_amount: totals.totalTaxable,
              total_tax_amount: totals.totalTax,
              grand_total: totals.grandTotal,
              line_items: validation.processedItems
          },
          validation_passed: true
      });
  });
});

// Enhanced save reviewed invoice endpoint with voucher type support
app.post('/api/save-reviewed-invoice', (req, res) => {
  const data = req.body;
  
  if (!data.id) {
      return res.status(400).json({ success: false, error: 'Invoice ID is required' });
  }
  
  console.log(`💾 Saving reviewed invoice ${data.id} with voucher type ${data.voucher_type}:`, data.line_items?.length || 0, 'items');
  
  // Handle line items if provided
  let lineItemsJson = '[]';
  let lineItemsCount = 0;
  let hasLineItems = false;
  let calculatedTaxableAmount = parseFloat(data.taxable_amount) || 0;
  let calculatedTotalTax = 0;
  let calculatedInvoiceValue = parseFloat(data.invoice_value) || 0;
  
  if (data.line_items && Array.isArray(data.line_items)) {
      // Validate and process line items
      const validation = LineItemsValidator.validateAndProcessLineItems(data.line_items);
      
      if (validation.isValid) {
          lineItemsJson = JSON.stringify(validation.processedItems);
          lineItemsCount = validation.processedItems.length;
          hasLineItems = lineItemsCount > 0;
          
          // Calculate totals from line items
          const totals = LineItemsValidator.calculateTotals(validation.processedItems);
          
          // Use line items totals if they exist, otherwise use provided values
          if (hasLineItems && totals.totalTaxable > 0) {
              calculatedTaxableAmount = totals.totalTaxable;
              calculatedTotalTax = totals.totalTax;
              
              // Recalculate invoice value including round off and TDS
              const roundOff = parseFloat(data.round_off) || 0;
              const tdsAmount = parseFloat(data.tds_amount) || 0;
              calculatedInvoiceValue = totals.totalTaxable + totals.totalTax + roundOff - tdsAmount;
          }
          
          console.log(`📊 Line items totals calculated:`, {
              taxable: calculatedTaxableAmount,
              tax: calculatedTotalTax,
              invoice: calculatedInvoiceValue
          });
      } else {
          return res.status(400).json({
              success: false,
              error: 'Line items validation failed',
              validationErrors: validation.errors
          });
      }
  }
  
  // Enhanced SQL update with voucher type support (preserve parseur_document_id)
  const sql = `UPDATE expenses SET 
      invoice_date = ?, invoice_number = ?, vendor_name = ?, vendor_gstn = ?, 
      customer_name = ?, customer_gstn = ?, place_of_supply = ?,
      taxable_amount = ?, igst_amount = ?, cgst_amount = ?, sgst_amount = ?, cess_amount = ?, round_off = ?, invoice_value = ?, 
      tds_rate = ?, tds_amount = ?, 
      description = ?, hsn_sac_code = ?, line_item_amount = ?, 
      quantity = ?, unit_rate = ?, 
      voucher_type = ?, ledger_name = ?, vendor_address = ?,
      total_tax_amount = ?, cgst_rate = ?, sgst_rate = ?, igst_rate = ?, cess_rate = ?,
      category = ?, status = ?,
      line_items = ?, line_items_count = ?, has_line_items = ?,
      table_structure_confidence = ?
      WHERE id = ?`;
  
  // Calculate total tax amount
  const totalTaxAmount = (parseFloat(data.igst_amount) || 0) + 
                        (parseFloat(data.cgst_amount) || 0) + 
                        (parseFloat(data.sgst_amount) || 0) + 
                        (parseFloat(data.cess_amount) || 0);
  
  const values = [
      // Basic invoice information
      prepareDateValue(data.invoice_date, 'invoice_date'), data.invoice_number, data.vendor_name, data.vendor_gstn,
      data.customer_name, data.customer_gstn, data.place_of_supply,
      
      // Amount information
      calculatedTaxableAmount, data.igst_amount, data.cgst_amount, data.sgst_amount,
      data.cess_amount, data.round_off, calculatedInvoiceValue,
      data.tds_rate, data.tds_amount,
      
      // Item information
      data.description, data.hsn_sac_code, data.line_item_amount,
      data.quantity, data.unit_rate,
      
      // Voucher type and additional information
      data.voucher_type || 'Purchase', data.ledger_name, data.vendor_address,
      
      // Tax information
      hasLineItems ? calculatedTotalTax : totalTaxAmount,
      data.cgst_rate, data.sgst_rate, data.igst_rate, data.cess_rate,
      
      // Status and category
      data.category || 'General', data.status || 'approved',
      
      // Line items information
      lineItemsJson, lineItemsCount, hasLineItems,
      hasLineItems ? 0.98 : 0, // High confidence for manually edited items
      
      // ID for WHERE clause
      data.id
  ];
  
  db.run(sql, values, function(err) {
      if (err) {
          console.error('❌ Save failed:', err);
          return res.status(500).json({ 
              success: false, 
              error: 'Failed to save invoice',
              details: err.message 
          });
      }
      
      if (this.changes === 0) {
          return res.status(404).json({ 
              success: false, 
              error: 'Invoice not found' 
          });
      }
      
      console.log(`✅ Invoice ${data.id} saved successfully with voucher type ${data.voucher_type}:`, {
          hasLineItems: hasLineItems,
          lineItemsCount: lineItemsCount,
          taxableAmount: calculatedTaxableAmount,
          invoiceValue: calculatedInvoiceValue,
          voucherType: data.voucher_type
      });
      
      res.json({
          success: true,
          message: 'Invoice updated successfully',
          id: data.id,
          voucher_type: data.voucher_type || 'Purchase',
          line_items_updated: hasLineItems,
          totals_recalculated: hasLineItems,
          final_amounts: {
              taxable_amount: calculatedTaxableAmount,
              total_tax_amount: hasLineItems ? calculatedTotalTax : totalTaxAmount,
              invoice_value: calculatedInvoiceValue,
              line_items_count: lineItemsCount
          }
      });
  });
});

// Export endpoint - ORIGINAL CSV FORMAT with voucher type
app.get('/api/export-tally', (req, res) => {
  const sql = 'SELECT * FROM expenses WHERE status = "approved" OR status = "processed" ORDER BY invoice_date DESC';
  
  db.all(sql, (err, rows) => {
      if (err) {
          console.error('Export query failed:', err);
          return res.status(500).json({ error: 'Failed to fetch data for export' });
      }
      
      // Enhanced headers for line items breakdown with voucher type
      const headers = [
          'Invoice Date', 'Invoice Number', 'Vendor Name', 'Vendor GSTN', 'Customer Name', 'Customer GSTN', 'Place of Supply',
          'Line Item Description', 'HSN/SAC Code', 'Quantity', 'Unit Rate', 'Line Item Amount', 'Tax Rate', 'Tax Amount',
          'Taxable Amount', 'IGST Amount', 'CGST Amount', 'SGST Amount', 'CESS Amount', 'Round Off', 'Invoice Value',
          'TDS Rate', 'TDS Amount', 'Total Tax Amount', 'Tax Type', 'Vendor State', 'Document Type', 'Status',
          'Line Items Count', 'Has Line Items', 'Voucher Type', 'Entry Type', 'Can Reprocess', 'AI Engine Document ID'
      ];
      
      let csv = headers.join(',') + '\n';
      let totalRows = 0;
      let totalInvoices = 0;
      
      rows.forEach(row => {
          totalInvoices++;
          const taxType = row.igst_amount > 0 ? 'Inter-state (IGST)' : 'Intra-state (CGST+SGST)';
          const vendorState = row.vendor_gstn ? GST_STATE_CODES[row.vendor_gstn.substring(0, 2)] || 'Unknown' : 'N/A';
          const canReprocess = !!(row.file_path || row.parseur_document_id);
          
          // Parse line items safely
          let lineItems = [];
          try {
              if (row.line_items && row.line_items !== '[]' && row.line_items !== 'null') {
                  const parsed = JSON.parse(row.line_items);
                  lineItems = Array.isArray(parsed) ? parsed : [];
              }
          } catch (e) {
              console.log('Failed to parse line items for invoice:', row.invoice_number);
              lineItems = [];
          }
          
          // Common invoice header data (REPEATS on every row)
          const invoiceHeaderData = [
              row.invoice_date || '',
              row.invoice_number || '',
              `"${(row.vendor_name || '').replace(/"/g, '""')}"`,
              row.vendor_gstn || '',
              `"${(row.customer_name || '').replace(/"/g, '""')}"`,
              row.customer_gstn || '',
              row.place_of_supply || ''
          ];
          
          // Invoice totals (REPEATS on every row)
          const invoiceTotals = [
              row.taxable_amount || 0,
              row.igst_amount || 0,
              row.cgst_amount || 0,
              row.sgst_amount || 0,
              row.cess_amount || 0,
              row.round_off || 0,
              row.invoice_value || 0,
              row.tds_rate || 0,
              row.tds_amount || 0,
              row.total_tax_amount || 0,
              taxType,
              vendorState,
              row.document_type || 'image',
              row.status || 'processed',
              row.line_items_count || 0,
              row.has_line_items || 0,
              row.voucher_type || 'Purchase',
              row.entry_type || 'upload',
              canReprocess ? 'Yes' : 'No',
              row.parseur_document_id || 'N/A'
          ];
          
          // Only keep line items that have an actual amount
          const meaningfulItems = lineItems.filter(item => (item.line_total || 0) > 0 || (item.unit_rate || 0) > 0);

          if (meaningfulItems.length > 0) {
              // Create one row per meaningful line item; invoice totals only on the first row
              meaningfulItems.forEach((item, index) => {
                  const lineItemData = [
                      `"${(item.description || `Line Item ${index + 1}`).replace(/"/g, '""')}"`,
                      item.hsn_code || item.hsn_sac_code || '',
                      item.quantity || 1,
                      item.unit_rate || 0,
                      item.line_total || 0,
                      item.tax_rate || 18,
                      item.tax_amount || 0
                  ];
                  const totalsForRow = index === 0 ? invoiceTotals : invoiceTotals.map(() => '');
                  const csvRow = [...invoiceHeaderData, ...lineItemData, ...totalsForRow];
                  csv += csvRow.join(',') + '\n';
                  totalRows++;
              });
          } else {
              // Single row for invoices without line items
              const lineItemData = [
                  `"${(row.description || 'Standard Invoice Item').replace(/"/g, '""')}"`,
                  row.hsn_sac_code || '',
                  row.quantity || 1,
                  row.unit_rate || 0,
                  row.line_item_amount || row.taxable_amount || 0,
                  18, // Default tax rate
                  row.total_tax_amount || 0
              ];
              
              const csvRow = [...invoiceHeaderData, ...lineItemData, ...invoiceTotals];
              csv += csvRow.join(',') + '\n';
              totalRows++;
          }
      });
      
      res.setHeader('Content-Type', 'text/csv');
      res.setHeader('Content-Disposition', 'attachment; filename=tally-import-with-line-items-and-reprocessing.csv');
      res.send(csv);
      
      console.log(`📊 Exported ${totalInvoices} invoices as ${totalRows} rows with enhanced duplicate detection and reprocessing info`);
  });
});

// 🚀 NEW: ENHANCED TALLY FORMAT EXPORT WITH AUTO-IMPORT FUNCTIONALITY AND FIXED LEDGER STANDARDIZATION TIMING
app.get('/api/export-tally-format', async (req, res) => {
  const sql = 'SELECT * FROM expenses WHERE status = "approved" OR status = "processed" ORDER BY invoice_date DESC';
  
  db.all(sql, async (err, rows) => {
      if (err) {
          console.error('Tally export query failed:', err);
          return res.status(500).json({ error: 'Failed to fetch data for Tally export' });
      }
      
      try {
          console.log('🚀 Starting enhanced Tally export with auto-import and fixed ledger standardization timing...');
          
          // 🎯 STEP 1: GENERATE EXCEL WITHOUT STANDARDIZATION (convertToTallyFormat generates dynamic names)
          console.log('📊 Converting data to Tally format (without standardization)...');
          let tallyData = convertToTallyFormat(rows);
          
          console.log(`✅ Excel data generated with ${tallyData.length} entries containing dynamic ledger names`);
          
          // Generate Excel file
          const wb = XLSX.utils.book_new();
          const ws = XLSX.utils.json_to_sheet(tallyData, { 
              header: [
                  "Voucher Date", "Voucher Type Name", "Voucher Number",
                  "Buyer/Supplier - Address", "Buyer/Supplier - Pincode",
                  "Ledger Name", "Ledger Amount", "Ledger Amount Dr/Cr",
                  "Item Name", "Billed Quantity", "Item Rate", "Item Rate per",
                  "Item Amount", "Change Mode ", "Voucher Narration"
              ]
          });
          
          XLSX.utils.book_append_sheet(wb, ws, "Accounting Voucher");
          
          // 🎯 STEP 2: SAVE EXCEL FILE ON SERVER
          const timestamp = new Date().toISOString().split('T')[0];
          const filename = `tally-vouchers-enhanced-${timestamp}.xlsx`;
          const serverFilePath = path.join(__dirname, 'tally-exports', filename);
          
          // Ensure tally-exports directory exists
          if (!fs.existsSync(path.join(__dirname, 'tally-exports'))) {
              fs.mkdirSync(path.join(__dirname, 'tally-exports'));
          }
          
          // Write Excel file to server
          XLSX.writeFile(wb, serverFilePath);
          console.log(`📁 Excel file saved to server: ${serverFilePath}`);
          
          // 🎯 STEP 3: LOAD LEDGER STANDARDIZATION FOR AUTO-IMPORT
          console.log('📚 Loading ledger standardization for auto-import...');
          await ledgerStandardizer.loadLedgerMaster();
          
          // 🎯 STEP 4: AUTO-IMPORT TO TALLY WITH PROPER LEDGER STANDARDIZATION TIMING
          let importResults = null;
          let importError = null;
          
          try {
              console.log('🔄 Starting auto-import to Tally with proper ledger standardization timing...');
              // 🎯 CRITICAL: Pass ledgerStandardizer to importToTally - standardization happens INSIDE the method
              importResults = await TallyIntegrator.importToTally(serverFilePath, ledgerStandardizer);
              console.log('✅ Tally auto-import completed successfully with proper standardization timing:', importResults);
          } catch (tallyError) {
              console.error('❌ Tally auto-import failed:', tallyError.message);
              importError = tallyError.message;
          }
          
          // 🎯 STEP 5: RESPOND WITH SUCCESS AND OPTIONAL DOWNLOAD
          const response = {
              success: true,
              message: `Enhanced Tally export with proper ledger standardization timing completed: ${tallyData.length} voucher entries from ${Object.keys(rows.reduce((groups, row) => { groups[row.invoice_number] = true; return groups; }, {})).length} invoices`,
              excel_generated: true,
              server_file_path: serverFilePath,
              filename: filename,
              voucher_entries_count: tallyData.length,
              invoices_processed: rows.length,
              ledger_standardization: {
                  timing_fixed: true,
                  applied_during_import: ledgerStandardizer.initialized,
                  excel_contains_dynamic_names: true,
                  standardization_happens_after_excel_load: true,
                  statistics: ledgerStandardizer.getStatistics(),
                  ledger_master_file: './Ledger extractionV2/TallyData_Complete_latest.xlsx',
                  sheet_used: '📋 Ledgers'
              },
              auto_import: {
                  attempted: true,
                  success: !importError,
                  error: importError,
                  results: importResults,
                  standardization_timing: 'fixed - applied after Excel load but before Tally processing'
              },
              download_available: true,
              download_url: `/api/download-tally-file/${filename}`
          };
          
          if (importResults) {
              response.tally_import_summary = {
                  total_rows: importResults.totalRows,
                  stock_items: `${importResults.stockItems.created} created, ${importResults.stockItems.exceptions} exceptions`,
                  ledgers: `${importResults.ledgers.created} created, ${importResults.ledgers.exceptions} exceptions`,
                  vouchers: `${importResults.vouchers.created} created, ${importResults.vouchers.exceptions} exceptions`,
                  errors: importResults.errors,
                  ledger_standardization: importResults.ledgerStandardization
              };
          }
          
          res.json(response);
          
      } catch (error) {
          console.error('❌ Enhanced Tally export failed:', error);
          res.status(500).json({ 
              success: false,
              error: 'Enhanced Tally export failed', 
              details: error.message 
          });
      }
  });
});

// 🆕 NEW: DOWNLOAD TALLY FILE ENDPOINT
app.get('/api/download-tally-file/:filename', (req, res) => {
  const filename = req.params.filename;
  const filePath = path.join(__dirname, 'tally-exports', filename);
  
  // Security check: ensure filename doesn't contain path traversal
  if (filename.includes('..') || filename.includes('/') || filename.includes('\\')) {
      return res.status(400).json({ error: 'Invalid filename' });
  }
  
  // Check if file exists
  if (!fs.existsSync(filePath)) {
      return res.status(404).json({ error: 'File not found' });
  }
  
  console.log(`📥 Downloading Tally file: ${filename}`);
  
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
  
  // Stream the file
 const fileStream = fs.createReadStream(filePath);
 fileStream.pipe(res);
});

// 🔐 ENHANCED BULK UPLOAD WITH CONTENT-BASED DUPLICATE DETECTION
app.post('/api/upload-bulk', upload.array('receipts', 20), async (req, res) => {
 try {
     if (!req.files || req.files.length === 0) {
         return res.status(400).json({ success: false, error: 'No files uploaded' });
     }
     
     const voucherType = req.body.voucher_type || 'Purchase';
     
     console.log('📦 Processing enhanced bulk upload:', req.files.length, 'files with voucher type:', voucherType);
     
     // 🔐 ENHANCED BULK DUPLICATE DETECTION
     const bulkValidation = await DuplicateDetector.validateBulkUpload(req.files);
     
     const results = {
         total: req.files.length,
         successful: [],
         failed: [],
         duplicates: bulkValidation.duplicates,
         conflicts: bulkValidation.conflicts,
         summary: { 
             total: req.files.length, 
             successful: 0, 
             failed: 0,
             duplicates: bulkValidation.batchStats.contentDuplicates,
             conflicts: bulkValidation.batchStats.nameConflicts,
             totalLineItems: 0,
             parseurProcessed: 0,
             documentAIProcessed: 0,
             reprocessableDocuments: 0,
             voucherType: voucherType
         }
     };
     
     // Add validation errors to failed results
     results.failed = bulkValidation.errors;
     results.summary.failed += bulkValidation.errors.length;
     
     // ── Shared DB insert helper (used for single-invoice and each invoice in multi-invoice PDFs) ──
     const insertExtractedInvoice = async (extractedData, file, validFile, result, voucherType, results) => {
         // Duplicate / continuation check
         const contentDuplicateCheck = await new Promise((resolve, reject) => {
             const hasInvoiceNumber = extractedData.invoice_number && extractedData.invoice_number !== '' && extractedData.invoice_number !== 'null';
             const hasGstn = extractedData.vendor_gstn && extractedData.vendor_gstn !== '' && extractedData.vendor_gstn !== 'null';
             let sql, params;
             if (hasInvoiceNumber) {
                 sql = `SELECT * FROM expenses WHERE invoice_number = ? AND invoice_number != '' AND invoice_number != 'null'`;
                 params = [extractedData.invoice_number];
             } else if (hasGstn) {
                 sql = `SELECT * FROM expenses WHERE vendor_gstn = ? AND vendor_gstn != '' AND vendor_gstn != 'null' ORDER BY id DESC LIMIT 1`;
                 params = [extractedData.vendor_gstn];
             } else {
                 return resolve(null);
             }
             db.get(sql, params, (err, row) => err ? reject(err) : resolve(row));
         });

         if (contentDuplicateCheck && extractedData.invoice_number) {
             const newItems = Array.isArray(extractedData.line_items) ? extractedData.line_items : [];
             const meaningfulNewItems = newItems.filter(i => i.description && i.description.trim().length > 2);
             const invoiceNumberMatches = contentDuplicateCheck.invoice_number === extractedData.invoice_number && extractedData.invoice_number !== '';
             const newInvoiceNumberMissing = !extractedData.invoice_number || extractedData.invoice_number === '' || extractedData.invoice_number === 'null';
             const gstnMatches = contentDuplicateCheck.vendor_gstn && extractedData.vendor_gstn && contentDuplicateCheck.vendor_gstn === extractedData.vendor_gstn;
             const isSameInvoice = invoiceNumberMatches || (gstnMatches && newInvoiceNumberMissing);

             if (isSameInvoice) {
                 let existingItems = [];
                 try { existingItems = JSON.parse(contentDuplicateCheck.line_items || '[]'); } catch (_) {}
                 const existingDescs = new Set(existingItems.map(i => (i.description || '').trim().toLowerCase()));
                 const trulyNew = meaningfulNewItems.filter(i => !existingDescs.has((i.description || '').trim().toLowerCase()));
                 const existingValue = parseFloat(contentDuplicateCheck.invoice_value) || 0;
                 const existingTaxable = parseFloat(contentDuplicateCheck.taxable_amount) || 0;
                 const existingCgst = parseFloat(contentDuplicateCheck.cgst_amount) || 0;
                 const existingSgst = parseFloat(contentDuplicateCheck.sgst_amount) || 0;
                 const existingIgst = parseFloat(contentDuplicateCheck.igst_amount) || 0;
                 const newValue = parseFloat(extractedData.invoice_value) || 0;
                 const newTaxable = parseFloat(extractedData.taxable_amount) || 0;
                 const newCgst = parseFloat(extractedData.cgst_amount) || 0;
                 const newSgst = parseFloat(extractedData.sgst_amount) || 0;
                 const newIgst = parseFloat(extractedData.igst_amount) || 0;
                 const mergedItems = trulyNew.length > 0 ? [...existingItems, ...trulyNew] : existingItems;
                 const finalValue = existingValue > 0 ? existingValue : newValue;
                 const finalTaxable = existingTaxable > 0 ? existingTaxable : newTaxable;
                 const finalCgst = existingCgst > 0 ? existingCgst : newCgst;
                 const finalSgst = existingSgst > 0 ? existingSgst : newSgst;
                 const finalIgst = existingIgst > 0 ? existingIgst : newIgst;
                 const totalsChanged = finalValue !== existingValue || finalTaxable !== existingTaxable || finalCgst !== existingCgst || finalSgst !== existingSgst || finalIgst !== existingIgst;
                 const itemsChanged = trulyNew.length > 0;
                 if (itemsChanged || totalsChanged) {
                     await new Promise((resolve, reject) => {
                         db.run(`UPDATE expenses SET line_items=?, line_items_count=?, has_line_items=1, invoice_value=?, taxable_amount=?, cgst_amount=?, sgst_amount=?, igst_amount=?, total_tax_amount=? WHERE id=?`,
                             [JSON.stringify(mergedItems), mergedItems.length, finalValue, finalTaxable, finalCgst, finalSgst, finalIgst, finalCgst + finalSgst + finalIgst, contentDuplicateCheck.id],
                             (err) => err ? reject(err) : resolve());
                     });
                     const what = [itemsChanged ? `${trulyNew.length} new line items` : '', totalsChanged ? 'totals' : ''].filter(Boolean).join(' + ');
                     console.log(`🔗 Merged ${what} into existing invoice ${extractedData.invoice_number} (ID: ${contentDuplicateCheck.id})`);
                     results.successful.push({ filename: file.originalname, invoice_number: extractedData.invoice_number, message: `Continuation page merged: ${what}` });
                     results.summary.successful++;
                 } else {
                     results.successful.push({ filename: file.originalname, invoice_number: extractedData.invoice_number, message: 'Continuation page already merged — no new data' });
                     results.summary.successful++;
                 }
                 return; // handled as continuation
             }

             // Different invoice numbers — genuine duplicate
             results.duplicates.push({ filename: file.originalname, error: 'Duplicate invoice detected', details: `Invoice ${extractedData.invoice_number || '(unknown)'} already exists` });
             results.summary.duplicates++;
             return;
         }

         // Fresh insert
         const insertSql = `INSERT INTO expenses (
             invoice_date, invoice_number, vendor_name, vendor_gstn, customer_name, customer_gstn, place_of_supply,
             taxable_amount, igst_amount, cgst_amount, sgst_amount, cess_amount, round_off, invoice_value,
             tds_rate, tds_amount, description, hsn_sac_code, line_item_amount, quantity, unit_rate,
             voucher_type, ledger_name, vendor_address, total_tax_amount,
             cgst_rate, sgst_rate, igst_rate, cess_rate, category, file_path, original_filename, file_hash,
             parseur_document_id, extracted_text,
             confidence_score, amount_confidence, document_type, processing_time_ms, status, processing_source,
             line_items, line_items_count, has_line_items, table_structure_confidence,
             processing_method, validation_errors, validation_warnings, entry_type
         ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`;
         const values = [
             prepareDateValue(extractedData.invoice_date, 'invoice_date'),
             prepareValue(extractedData.invoice_number, 'invoice_number'),
             prepareValue(extractedData.vendor_name, 'vendor_name') || 'Unknown Vendor',
             prepareValue(extractedData.vendor_gstn, 'vendor_gstn'),
             prepareValue(extractedData.customer_name, 'customer_name'),
             prepareValue(extractedData.customer_gstn, 'customer_gstn'),
             prepareValue(extractedData.place_of_supply, 'place_of_supply'),
             prepareNumericValue(extractedData.taxable_amount || extractedData.taxable_value, 'taxable_amount'),
             prepareNumericValue(extractedData.igst_amount, 'igst_amount'),
             prepareNumericValue(extractedData.cgst_amount, 'cgst_amount'),
             prepareNumericValue(extractedData.sgst_amount, 'sgst_amount'),
             prepareNumericValue(extractedData.cess_amount, 'cess_amount'),
             prepareNumericValue(extractedData.round_off, 'round_off'),
             prepareNumericValue(extractedData.invoice_value || extractedData.amount, 'invoice_value'),
             prepareNumericValue(extractedData.tds_rate, 'tds_rate'),
             prepareNumericValue(extractedData.tds_amount, 'tds_amount'),
             prepareValue(extractedData.description, 'description'),
             prepareValue(extractedData.hsn_sac_code, 'hsn_sac_code'),
             prepareNumericValue(extractedData.line_item_amount, 'line_item_amount'),
             prepareNumericValue(extractedData.quantity, 'quantity'),
             prepareNumericValue(extractedData.unit_rate, 'unit_rate'),
             prepareValue(extractedData.voucher_type, 'voucher_type') || voucherType,
             prepareValue(extractedData.ledger_name, 'ledger_name'),
             prepareValue(extractedData.vendor_address, 'vendor_address'),
             prepareNumericValue(extractedData.total_tax_amount, 'total_tax_amount'),
             prepareNumericValue(extractedData.cgst_rate, 'cgst_rate'),
             prepareNumericValue(extractedData.sgst_rate, 'sgst_rate'),
             prepareNumericValue(extractedData.igst_rate, 'igst_rate'),
             prepareNumericValue(extractedData.cess_rate, 'cess_rate'),
             prepareValue(extractedData.category, 'category') || 'General',
             file.path,
             file.originalname,
             validFile.validation.fileHash,
             prepareValue(extractedData.parseur_document_id, 'parseur_document_id'),
             prepareValue(result.rawText, 'extracted_text'),
             prepareNumericValue(extractedData.confidence_score, 'confidence_score', 0.8),
             prepareNumericValue(extractedData.amount_confidence, 'amount_confidence', 0.8),
             path.extname(file.originalname).toLowerCase() === '.pdf' ? 'pdf' : 'image',
             prepareNumericValue(extractedData.processing_time_ms, 'processing_time_ms', 0),
             'pending_review',
             prepareValue(extractedData.processing_source || result.processingSource, 'processing_source'),
             JSON.stringify(extractedData.line_items || []),
             prepareNumericValue(extractedData.line_items_count, 'line_items_count', 0),
             prepareBooleanValue(extractedData.has_line_items, 'has_line_items'),
             prepareNumericValue(extractedData.table_structure_confidence, 'table_structure_confidence', 0),
             prepareValue(extractedData.processing_method, 'processing_method') || 'auto',
             prepareValue(extractedData.validation_errors, 'validation_errors') || '[]',
             prepareValue(extractedData.validation_warnings, 'validation_warnings') || '[]',
             'bulk_upload'
         ];
         await new Promise((resolve, reject) => { db.run(insertSql, values, function(err) { if (err) reject(err); else resolve(this.lastID); }); });
         results.successful.push({
             filename: file.originalname,
             expense: {
                 invoice_number: extractedData.invoice_number,
                 vendor_name: extractedData.vendor_name,
                 invoice_value: extractedData.invoice_value || extractedData.amount,
                 line_items_count: extractedData.line_items_count || 0,
                 processing_source: result.processingSource,
                 voucher_type: extractedData.voucher_type,
                 can_reprocess: !!(extractedData.parseur_document_id),
                 parseur_document_id: extractedData.parseur_document_id || null
             }
         });
         results.summary.totalLineItems += extractedData.line_items_count || 0;
         if (result.processingSource === 'parseur' || result.processingSource === 'parseur_reprocessed') results.summary.parseurProcessed++;
         else if (result.processingSource === 'document_ai') results.summary.documentAIProcessed++;
         if (extractedData.parseur_document_id) results.summary.reprocessableDocuments++;
     };

     // ── Phase 1: Run all Claude AI extractions IN PARALLEL ──────────────────────────────────
     // This reduces total wait from N×30s to ~30s for any batch size.
     // DB operations remain sequential below to avoid SQLite write contention.
     console.log(`⚡ Starting parallel AI extraction for ${bulkValidation.valid.length} file(s)...`);
     const extractionResults = await Promise.allSettled(
         bulkValidation.valid.map(async (validFile) => {
             const file = req.files.find(f => f.originalname === validFile.filename);
             if (!file) return { skipped: true, validFile };
             const result = await UnifiedDocumentProcessor.processInvoice(file.path, file.originalname, voucherType);
             return { file, validFile, result };
         })
     );
     console.log(`✅ Parallel extraction complete — processing DB operations...`);

     // ── Phase 2: Sequential DB operations for each extraction result ─────────────────────
     for (const settled of extractionResults) {
         let file, validFile, result;
         try {
             if (settled.status === 'rejected') {
                 // AI extraction itself failed — push to failed list; count reconciled at end
                 results.failed.push({ filename: '(unknown)', error: 'Processing failed', details: settled.reason?.message || String(settled.reason) });
                 continue;
             }
             const val = settled.value;
             if (!val || val.skipped) continue;
             ({ file, validFile, result } = val);

             // ── Multi-invoice PDF: expand into individual invoices and process each ───────
             if (result.multipleInvoices && result.multipleInvoices.length > 1) {
                 console.log(`📄 Expanding ${result.multipleInvoices.length} invoices from ${file.originalname}`);
                 for (const invData of result.multipleInvoices) {
                     invData.voucher_type = voucherType;
                     invData.entry_type = 'bulk_upload';
                     // Reuse the same DB insert helper by temporarily overriding result.extractedData
                     await insertExtractedInvoice(invData, file, validFile, result, voucherType, results);
                 }
                 continue; // skip the single-invoice path below
             }

             const extractedData = result.extractedData;
             extractedData.voucher_type = voucherType;
             extractedData.entry_type = 'bulk_upload';
             await insertExtractedInvoice(extractedData, file, validFile, result, voucherType, results);
             
         } catch (error) {
             // Clean up file on error
             if (file && file.path) {
                 await DuplicateDetector.safeDeleteFile(file.path);
             }
             results.failed.push({
                 filename: validFile ? validFile.filename : (file ? file.originalname : '(unknown)'),
                 error: 'Processing failed',
                 details: error.message
             });
         }
     }
     
     results.summary.successful = results.successful.length;
     results.summary.failed += results.failed.length - bulkValidation.errors.length; // Don't double count validation errors
     
     console.log('✅ Enhanced bulk upload completed with content-based duplicate detection:', results.summary);
     
     res.json({
         success: true,
         message: `Enhanced bulk upload completed: ${results.summary.successful} successful, ${results.summary.failed} failed, ${results.summary.duplicates} duplicates, ${results.summary.conflicts} conflicts`,
         ...results,
         duplicate_detection: {
             method: 'SHA-256 content-based',
             batch_validation: 'enabled',
             reprocessing_support: 'enabled'
         }
     });
     
 } catch (error) {
     console.error('❌ Enhanced bulk upload failed:', error);
     res.status(500).json({
         success: false,
         error: 'Enhanced bulk upload failed',
         details: error.message
     });
 }
});

// Approval endpoints
app.post('/api/expenses/:id/approve', (req, res) => {
 const expenseId = req.params.id;
 const sql = 'UPDATE expenses SET status = ? WHERE id = ?';

 db.run(sql, ['approved', expenseId], function(err) {
     if (err) {
         console.error('Approval failed:', err);
         return res.status(500).json({ success: false, error: 'Failed to approve expense' });
     }
     
     if (this.changes === 0) {
         return res.status(404).json({ success: false, error: 'Expense not found' });
     }
     
     res.json({
         success: true,
         message: 'Expense approved successfully',
         id: expenseId,
         status: 'approved'
     });
 });
});

app.post('/api/expenses/approve-all', (req, res) => {
 const { ids } = req.body;

 if (!ids || !Array.isArray(ids) || ids.length === 0) {
     return res.status(400).json({ success: false, error: 'Invalid IDs array' });
 }

 const placeholders = ids.map(() => '?').join(',');
 const sql = `UPDATE expenses SET status = 'approved' WHERE id IN (${placeholders})`;

 db.run(sql, ids, function(err) {
     if (err) {
         console.error('Bulk approval failed:', err);
         return res.status(500).json({ success: false, error: 'Failed to approve expenses' });
     }
     
     res.json({
         success: true,
         message: `${this.changes} expenses approved successfully`,
         approved_count: this.changes,
         ids: ids
     });
 });
});

// Delete endpoint for expenses
app.delete('/api/expenses/:id', (req, res) => {
 const expenseId = req.params.id;
 const sql = 'DELETE FROM expenses WHERE id = ?';
 
 db.run(sql, [expenseId], function(err) {
     if (err) {
         console.error('Delete failed:', err);
         return res.status(500).json({ success: false, error: 'Failed to delete expense' });
     }
     
     if (this.changes === 0) {
         return res.status(404).json({ success: false, error: 'Expense not found' });
     }
     
     res.json({
         success: true,
         message: 'Expense deleted successfully',
         id: expenseId
     });
 });
});

// Get line items for a specific invoice
app.get('/api/expenses/:id/line-items', (req, res) => {
 const expenseId = req.params.id;
 const sql = 'SELECT line_items, line_items_count, has_line_items, table_structure_confidence, processing_source, voucher_type, parseur_document_id, taxable_amount, igst_amount, cgst_amount, sgst_amount, invoice_value FROM expenses WHERE id = ?';

 db.get(sql, [expenseId], (err, row) => {
     if (err) {
         console.error('Line items query failed:', err);
         return res.status(500).json({ error: 'Failed to fetch line items' });
     }
     
     if (!row) {
         return res.status(404).json({ error: 'Invoice not found' });
     }
     
     try {
         const lineItems = row.line_items ? JSON.parse(row.line_items) : [];
         
         res.json({
             success: true,
             line_items: lineItems,
             line_items_count: row.line_items_count || 0,
             has_line_items: row.has_line_items || false,
             table_structure_confidence: row.table_structure_confidence || 0,
             processing_source: row.processing_source || 'unknown',
             voucher_type: row.voucher_type || 'Purchase',
             can_reprocess: !!(row.parseur_document_id),
             parseur_document_id: row.parseur_document_id || null,
             invoice_totals: {
                 taxable_amount: row.taxable_amount || 0,
                 igst_amount: row.igst_amount || 0,
                 cgst_amount: row.cgst_amount || 0,
                 sgst_amount: row.sgst_amount || 0,
                 invoice_value: row.invoice_value || 0
             }
         });
     } catch (parseError) {
         res.status(500).json({ error: 'Failed to parse line items data' });
     }
 });
});

// 🆕 NEW LEDGER STANDARDIZATION ANALYTICS ENDPOINT
app.get('/api/analytics/ledger-standardization', (req, res) => {
 res.json(ledgerStandardizer.getStatistics());
});

// 🔐 NEW DUPLICATE DETECTION ANALYTICS ENDPOINT
app.get('/api/analytics/duplicates', (req, res) => {
 const sql = `SELECT 
     COUNT(*) as totalUploads,
     COUNT(DISTINCT original_filename) as uniqueFilenames,
     COUNT(DISTINCT file_hash) as uniqueFileHashes,
     COUNT(*) - COUNT(DISTINCT original_filename) as filenameCollisions,
     COUNT(*) - COUNT(DISTINCT file_hash) as contentDuplicates,
     COUNT(CASE WHEN voucher_type = 'Purchase' THEN 1 END) as purchaseUploads,
     COUNT(CASE WHEN voucher_type = 'Journal' THEN 1 END) as journalUploads,
     COUNT(CASE WHEN entry_type = 'manual' THEN 1 END) as manualEntries,
     COUNT(CASE WHEN entry_type = 'payment_journal' THEN 1 END) as paymentJournalEntries,
     COUNT(CASE WHEN parseur_document_id IS NOT NULL THEN 1 END) as reprocessableDocuments,
     COUNT(CASE WHEN processing_source = 'parseur_reprocessed' THEN 1 END) as reprocessedDocuments
 FROM expenses`;

 db.get(sql, (err, stats) => {
     if (err) {
         console.error('Duplicate analytics query failed:', err);
         return res.status(500).json({ error: 'Failed to fetch duplicate analytics' });
     }
     
     res.json({
         totalUploads: stats.totalUploads || 0,
         uniqueFilenames: stats.uniqueFilenames || 0,
         uniqueFileHashes: stats.uniqueFileHashes || 0,
         filenameCollisions: stats.filenameCollisions || 0,
         contentDuplicates: stats.contentDuplicates || 0,
         duplicateDetectionRate: stats.totalUploads > 0 ? 
             Math.round(((stats.filenameCollisions + stats.contentDuplicates) / stats.totalUploads) * 100) : 0,
         voucherTypeBreakdown: {
             purchase: stats.purchaseUploads || 0,
             journal: stats.journalUploads || 0
         },
         entryTypeBreakdown: {
             manual: stats.manualEntries || 0,
             paymentJournal: stats.paymentJournalEntries || 0,
             uploaded: (stats.totalUploads || 0) - (stats.manualEntries || 0) - (stats.paymentJournalEntries || 0)
         },
         // 🧠 Reprocessing analytics
         reprocessableDocuments: stats.reprocessableDocuments || 0,
         reprocessedDocuments: stats.reprocessedDocuments || 0,
         reprocessingAvailabilityRate: stats.totalUploads > 0 ? 
             Math.round((stats.reprocessableDocuments / stats.totalUploads) * 100) : 0,
         reprocessingUsageRate: stats.reprocessableDocuments > 0 ? 
             Math.round((stats.reprocessedDocuments / stats.reprocessableDocuments) * 100) : 0,
         duplicateDetectionMethod: 'SHA-256 content-based hashing',
         creditSavings: {
             duplicatesBlocked: stats.contentDuplicates || 0,
             estimatedCreditsSaved: (stats.contentDuplicates || 0) * 1, // Assuming 1 credit per document
             reprocessingEvents: stats.reprocessedDocuments || 0
         }
     });
 });
});

// Analytics endpoint for line items with voucher type breakdown
app.get('/api/analytics/line-items', (req, res) => {
 const sql = `SELECT 
     COUNT(*) as totalInvoices,
     COUNT(CASE WHEN has_line_items = 1 THEN 1 END) as invoicesWithLineItems,
     SUM(line_items_count) as totalLineItems,
     AVG(line_items_count) as avgLineItemsPerInvoice,
     AVG(table_structure_confidence) as avgTableConfidence,
     COUNT(CASE WHEN line_items_count = 1 THEN 1 END) as singleItemInvoices,
     COUNT(CASE WHEN line_items_count BETWEEN 2 AND 5 THEN 1 END) as smallInvoices,
     COUNT(CASE WHEN line_items_count BETWEEN 6 AND 10 THEN 1 END) as mediumInvoices,
     COUNT(CASE WHEN line_items_count > 10 THEN 1 END) as largeInvoices,
     COUNT(CASE WHEN processing_source = 'parseur' THEN 1 END) as parseurProcessed,
     COUNT(CASE WHEN processing_source = 'parseur_reprocessed' THEN 1 END) as parseurReprocessed,
     COUNT(CASE WHEN processing_source = 'document_ai' THEN 1 END) as documentAIProcessed,
     COUNT(CASE WHEN voucher_type = 'Purchase' THEN 1 END) as purchaseVouchers,
     COUNT(CASE WHEN voucher_type = 'Journal' THEN 1 END) as journalVouchers,
     COUNT(CASE WHEN entry_type = 'manual' THEN 1 END) as manualEntries,
     COUNT(CASE WHEN entry_type = 'payment_journal' THEN 1 END) as paymentJournalEntries,
     COUNT(CASE WHEN parseur_document_id IS NOT NULL THEN 1 END) as reprocessableDocuments,
     MAX(line_items_count) as maxLineItems,
     MIN(line_items_count) as minLineItems
 FROM expenses WHERE status = 'approved' OR status = 'processed'`;

 db.get(sql, (err, stats) => {
     if (err) {
         console.error('Line items analytics query failed:', err);
         return res.status(500).json({ error: 'Failed to fetch line items analytics' });
     }
     
     res.json({
         totalInvoices: stats.totalInvoices || 0,
         invoicesWithLineItems: stats.invoicesWithLineItems || 0,
         totalLineItems: stats.totalLineItems || 0,
         avgLineItemsPerInvoice: Math.round((stats.avgLineItemsPerInvoice || 0) * 10) / 10,
         avgTableConfidence: Math.round((stats.avgTableConfidence || 0) * 100),
         distribution: {
             singleItem: stats.singleItemInvoices || 0,
             small: stats.smallInvoices || 0,
             medium: stats.mediumInvoices || 0,
             large: stats.largeInvoices || 0
         },
         range: {
             max: stats.maxLineItems || 0,
             min: stats.minLineItems || 0
         },
         processing: {
             parseur: stats.parseurProcessed || 0,
             parseurReprocessed: stats.parseurReprocessed || 0,
             documentAI: stats.documentAIProcessed || 0
         },
         voucherTypes: {
             purchase: stats.purchaseVouchers || 0,
             journal: stats.journalVouchers || 0
         },
         entryTypes: {
             manual: stats.manualEntries || 0,
             paymentJournal: stats.paymentJournalEntries || 0,
             uploaded: (stats.totalInvoices || 0) - (stats.manualEntries || 0) - (stats.paymentJournalEntries || 0)
         },
         reprocessing: {
             available: stats.reprocessableDocuments || 0,
             used: stats.parseurReprocessed || 0,
             availabilityRate: stats.totalInvoices > 0 ? 
                 Math.round((stats.reprocessableDocuments / stats.totalInvoices) * 100) : 0
         },
         lineItemsExtractionRate: stats.totalInvoices > 0 ? 
             Math.round((stats.invoicesWithLineItems / stats.totalInvoices) * 100) : 0,
         parseurDataPreservationRate: stats.totalInvoices > 0 ? 
             Math.round(((stats.parseurProcessed + stats.parseurReprocessed) / stats.totalInvoices) * 100) : 0
     });
 });
});

// Dashboard Statistics Endpoint with Payment Journal Support
app.get('/api/dashboard-stats', (req, res) => {
 const sql = `SELECT 
     COUNT(*) as totalCount,
     SUM(invoice_value) as totalExpenses,
     AVG(confidence_score) as avgConfidence,
     SUM(line_items_count) as totalLineItems,
     COUNT(CASE WHEN processing_source = 'parseur' OR processing_source = 'parseur_reprocessed' THEN 1 END) as parseurProcessed,
     COUNT(CASE WHEN processing_source = 'document_ai' THEN 1 END) as documentAIFallback,
     COUNT(CASE WHEN has_line_items = 1 THEN 1 END) as invoicesWithLineItems,
     AVG(CASE WHEN line_items_count > 0 THEN line_items_count ELSE NULL END) as avgLineItemsPerInvoice,
     COUNT(CASE WHEN voucher_type = 'Purchase' THEN 1 END) as purchaseVouchers,
     COUNT(CASE WHEN voucher_type = 'Journal' THEN 1 END) as journalVouchers,
     COUNT(CASE WHEN entry_type = 'manual' THEN 1 END) as manualEntries,
     COUNT(CASE WHEN entry_type = 'payment_journal' THEN 1 END) as paymentJournalEntries,
     SUM(cgst_amount) as totalCGST,
     SUM(sgst_amount) as totalSGST,
     SUM(igst_amount) as totalIGST,
     SUM(total_tax_amount) as totalTaxes
 FROM expenses`;

 db.get(sql, (err, stats) => {
     if (err) {
         console.error('Dashboard stats query failed:', err);
         return res.status(500).json({ error: 'Failed to fetch dashboard statistics' });
     }
     
     // Format the response with proper defaults
     const response = {
         totalCount: stats.totalCount || 0,
         totalExpenses: stats.totalExpenses || 0,
         avgConfidence: stats.avgConfidence ? (stats.avgConfidence * 100) : 0, // Convert to percentage
         totalLineItems: stats.totalLineItems || 0,
         parseurProcessed: stats.parseurProcessed || 0,
         documentAIFallback: stats.documentAIFallback || 0,
         invoicesWithLineItems: stats.invoicesWithLineItems || 0,
         avgLineItemsPerInvoice: Math.round((stats.avgLineItemsPerInvoice || 0) * 10) / 10,
         purchaseVouchers: stats.purchaseVouchers || 0,
         journalVouchers: stats.journalVouchers || 0,
         manualEntries: stats.manualEntries || 0,
         paymentJournalEntries: stats.paymentJournalEntries || 0,
         totalCGST: stats.totalCGST || 0,
         totalSGST: stats.totalSGST || 0,
         totalIGST: stats.totalIGST || 0,
         totalTaxes: stats.totalTaxes || 0
     };
     
     console.log('📊 Dashboard stats retrieved:', {
         invoices: response.totalCount,
         amount: response.totalExpenses,
         confidence: Math.round(response.avgConfidence) + '%',
         lineItems: response.totalLineItems,
         journals: response.journalVouchers,
         paymentJournals: response.paymentJournalEntries
     });
     
     res.json(response);
 });
});

// Global multer error handler — catches "Unexpected field" and other multer errors
// Must be registered BEFORE app.listen and AFTER all routes
app.use((err, req, res, next) => {
    if (err && err.code && err.code.startsWith('LIMIT_')) {
        // Multer error
        console.error(`❌ Multer error on ${req.method} ${req.path}: ${err.code} — field: ${err.field || 'unknown'}`);
        return res.status(400).json({ success: false, error: `File upload error: ${err.message}` });
    }
    // Other errors
    console.error(`❌ Unhandled error on ${req.method} ${req.path}:`, err.message);
    res.status(500).json({ success: false, error: err.message });
});

// Initialize database and start server
setupDatabase();
setTimeout(initializePdfConverter, 1000);

const httpServer = app.listen(PORT, () => {
 console.log('🚀 Simplifier - Enhanced Invoice Processing System with Fixed Ledger Standardization Timing');
 console.log(`📊 Server running on http://localhost:${PORT}`);
 console.log('');
 console.log('🎯 FIXED LEDGER STANDARDIZATION TIMING:');
 console.log('  • ✅ Excel Generation: Dynamic names (e.g., "Input CGST @ 9%") created without standardization');
 console.log('  • ✅ Excel Loading: Data loaded from file in TallyIntegrator.importToTally()');
 console.log('  • ✅ Standardization Applied: AFTER loading Excel but BEFORE processing for Tally');
 console.log('  • ✅ Tally Import: Uses standardized names (e.g., "CGST INPUT") from master file');
 console.log('  • ✅ Master File: ./Ledger extractionV2/TallyData_Complete_latest.xlsx');
 console.log('  • ✅ Sheet: "📋 Ledgers" with "Ledger Name" and "Aliases" columns');
 console.log('');
 console.log('🔧 TIMING FIX DETAILS:');
 console.log('  1. convertToTallyFormat() - Generates dynamic ledger names (NO standardization)');
 console.log('  2. Excel saved to server with dynamic names');
 console.log('  3. TallyIntegrator.importToTally() - Loads Excel data');
 console.log('  4. 🎯 STANDARDIZATION APPLIED HERE - After load, before processing');
 console.log('  5. Process standardized data for Tally import');
 console.log('  6. Import to Tally with correct ledger names');
 console.log('');
 console.log('🎯 LEDGER STANDARDIZATION WORKFLOW (FIXED):');
 console.log('  • "Input CGST @ 9%" (Excel) → Load → Standardize → "CGST INPUT" (Tally)');
 console.log('  • "Input SGST @ 9%" (Excel) → Load → Standardize → "SGST INPUT" (Tally)');
 console.log('  • "Input IGST @ 18%" (Excel) → Load → Standardize → "IGST INPUT" (Tally)');
 console.log('  • Any alias in master file will be standardized to official name');
 console.log('');
 console.log('🎯 INTEGRATED TALLY AUTO-IMPORT FEATURES:');
 console.log('  • Click "Export Tally" button → Excel saved to server → Auto-imported to Tally');
 console.log('  • Seamless integration: No manual steps required');
 console.log('  • Optional download: Excel file available for download if needed');
 console.log('  • Real-time feedback: Shows import success/failure with detailed results');
 console.log('  • Enhanced error handling: Detailed import logs and troubleshooting');
 console.log('  • 🎯 FIXED: Ledger standardization happens at the correct time during import');
 console.log('');
 console.log('🔐 ENHANCED DUPLICATE DETECTION FEATURES:');
 console.log('  • SHA-256 content-based file hashing for accurate duplicate detection');
 console.log('  • Filename collision detection with content verification');
 console.log('  • Automatic cleanup of duplicate files to save disk space');
 console.log('  • Two-level validation: filename and content hash verification');
 console.log('  • Bulk upload optimization with batch duplicate detection');
 console.log('');
 console.log('🧠 PARSEUR DOCUMENT ID REPROCESSING FEATURES:');
 console.log('  • Store Parseur document IDs for all processed invoices');
 console.log('  • POST /api/expenses/:id/reprocess - Reprocess using existing Parseur data');
 console.log('  • Avoid unnecessary API calls and credit usage on duplicates');
 console.log('  • Reprocessing availability detection and user guidance');
 console.log('  • Fast reprocessing with existing document data (500ms vs full processing)');
 console.log('');
 console.log('💰 PAYMENT JOURNAL VOUCHER FEATURES:');
 console.log('  • POST /api/payment-entry - Create Journal vouchers for direct payments');
 console.log('  • Debit/Credit entry validation with automatic balancing checks');
 console.log('  • Multi-ledger support for complex payment transactions');
 console.log('  • Enhanced Tally export with proper Journal voucher formatting');
 console.log('  • Unified storage in expenses table with voucher_type = "Journal"');
 console.log('  • Complete audit trail and line items support for journal entries');
 console.log('');
 console.log('🚀 ENHANCED ENDPOINTS WITH FIXED LEDGER STANDARDIZATION:');
 console.log('  • GET /api/export-tally-format - Fixed timing: standardization after Excel load');
 console.log('  • GET /api/download-tally-file/:filename - Download generated Excel files');
 console.log('  • POST /api/payment-entry - Create payment journal vouchers');
 console.log('  • POST /api/expenses/:id/reprocess - Reprocess with existing Parseur document ID');
 console.log('  • GET /api/analytics/duplicates - Duplicate detection analytics and credit savings');
 console.log('  • GET /api/analytics/ledger-standardization - Ledger standardization statistics');
 console.log('  • Enhanced /api/upload-invoice with content-based duplicate detection');
 console.log('  • Enhanced /api/upload-bulk with batch duplicate validation');
 console.log('  • Enhanced /api/expenses with reprocessing capability indicators');
 console.log('  • Enhanced dashboard stats with payment journal support');
 console.log('');
 console.log('📊 DATABASE ENHANCEMENTS:');
 console.log('  • Added parseur_document_id column for reprocessing support');
 console.log('  • Enhanced file_hash storage with SHA-256 content hashing');
 console.log('  • Duplicate detection metadata and analytics tracking');
 console.log('  • Processing source tracking for reprocessed documents');
 console.log('  • Journal voucher data storage in extracted_text field');
 console.log('  • Entry type tracking (upload, manual, payment_journal, bulk_upload, reprocessed)');
 console.log('');
 console.log('🔧 ENHANCED FUNCTIONALITY:');
 console.log('  • Content-based duplicate detection prevents unnecessary processing');
 console.log('  • Parseur credit savings through intelligent reprocessing');
 console.log('  • Enhanced bulk upload with sophisticated validation');
 console.log('  • Export functionality includes reprocessing metadata and journal vouchers');
 console.log('  • Analytics track duplicate detection effectiveness and journal entries');
 console.log('  • Payment tab for direct journal voucher creation');
 console.log('  • Debit/Credit entry validation with balance verification');
 console.log('  • Multi-ledger payment processing with narration support');
 console.log('  • Enhanced Tally export with Journal voucher formatting');
 console.log('  • 🎯 FIXED: Ledger standardization timing ensures correct alias mapping');
 console.log('');
 
 if (CLAUDE_CONFIG.enabled) {
     console.log('✅ CLAUDE AI PROCESSING ENABLED (PRIMARY EXTRACTOR):');
     console.log('   • Model: claude-sonnet-4-6 with vision capabilities');
     console.log('   • Supports PDF and image invoices (JPG, PNG, WEBP, GIF)');
     console.log('   • GSTIN format validation with state-code matching');
     console.log('   • Full line-item extraction with HSN/SAC codes');
     console.log('   • IGST / CGST+SGST / CESS / TDS extraction');
     console.log('   • File-based reprocessing without re-upload');
 } else {
     console.log('⚠️  CLAUDE AI DISABLED — set ANTHROPIC_API_KEY in .env to enable as primary extractor');
 }
 if (PARSEUR_CONFIG.enabled) {
     console.log('✅ PARSEUR configured as first fallback extractor');
 }
 
 console.log('');
 console.log('🎯 TALLY INTEGRATION STATUS:');
 console.log('   • Tally URL: http://localhost:9000 (ensure Tally is running with API enabled)');
 console.log('   • Auto-import: Enabled (triggers on Export Tally button click)');
 console.log('   • Excel storage: ./tally-exports/ directory');
 console.log('   • Stock items: Auto-created if needed');
 console.log('   • Ledgers: Auto-created with smart categorization & standardization');
 console.log('   • Vouchers: Created with inventory allocations');
 console.log('   • Journal entries: Proper debit/credit formatting');
 console.log('   • 🎯 FIXED: Ledger standardization timing ensures aliases work correctly');
 console.log('');
 console.log('📚 LEDGER STANDARDIZATION REQUIREMENTS (FIXED TIMING):');
 console.log('   • File path: ./Ledger extractionV2/TallyData_Complete_latest.xlsx');
 console.log('   • Sheet name: "📋 Ledgers"');
 console.log('   • Required columns: "Ledger Name" and "Aliases"');
 console.log('   • Aliases format: Comma-separated values (e.g., "Input CGST @ 9%, CGST Input")');
 console.log('   • Matching: Case-insensitive, trimmed exact matching');
 console.log('   • 🎯 TIMING: Applied after Excel load but before Tally processing');
 console.log('   • Fallback: Original names used if no match found');
 console.log('');
 console.log('🎉 COMPLETE ENHANCED FUNCTIONALITY WITH FIXED LEDGER STANDARDIZATION TIMING:');
 console.log('   🔐 Content-based duplicate detection with SHA-256 hashing');
 console.log('   🧠 Parseur document ID reprocessing for credit savings');
 console.log('   💰 Payment Journal voucher creation and processing');
 console.log('   📋 Enhanced CSV export with reprocessing metadata');
 console.log('   📊 🚀 AUTO-IMPORT Excel export (Tally vouchers) with fixed standardization timing');
 console.log('   📚 🎯 LEDGER STANDARDIZATION: Fixed timing - works with ANY aliases in master file');
 console.log('   ✏️ Manual entry support (no reprocessing available)');
 console.log('   🔄 Intelligent reprocessing with existing Parseur data');
 console.log('   💸 Credit optimization through duplicate prevention');
 console.log('   ⚡ Multi-tab structure with enhanced duplicate handling and Payment tab');
 console.log('   🔧 Complete line items editing with reprocessing support');
 console.log('   📈 Analytics for duplicate detection, reprocessing usage, and journal entries');
 console.log('   🎯 ONE-CLICK TALLY IMPORT: Export button now saves + imports with fixed standardization!');
 console.log('   📚 ALIAS MAPPING: Fixed timing ensures "Input CGST @ 9%" → "CGST INPUT" works!');
 console.log('');
 console.log('🔒 SECURITY & OPTIMIZATION FEATURES:');
 console.log('   • SHA-256 file hash-based duplicate detection (collision-resistant)');
 console.log('   • Automatic cleanup of duplicate files (storage optimization)');
 console.log('   • Content verification before processing (prevents false uploads)');
 console.log('   • Parseur document ID storage and reuse (credit optimization)');
 console.log('   • Batch validation for bulk uploads (performance optimization)');
 console.log('   • Comprehensive analytics and monitoring (usage tracking)');
 console.log('   • Filename conflict detection with content verification');
 console.log('   • Credit usage optimization through intelligent reprocessing');
 console.log('   • Journal voucher balance validation and error prevention');
 console.log('   • Secure file download with path traversal protection');
 console.log('   • 🎯 FIXED: Ledger name standardization timing prevents incorrect ledger creation');
 console.log('');
 console.log('💡 LEDGER STANDARDIZATION TIMING FIX:');
 console.log('   ❌ OLD: Standardization during Excel generation (too early)');
 console.log('   ❌ OLD: Standardization during export endpoint (wrong place)');
 console.log('   ✅ NEW: Standardization in TallyIntegrator.importToTally() after Excel load');
 console.log('   ✅ NEW: Dynamic names in Excel, standardized names in Tally');
 console.log('   ✅ NEW: Works with ANY aliases defined in master file');
 console.log('');
 console.log('🎯 FIXED WORKFLOW:');
 console.log('   1. Generate Excel with "Input CGST @ 9%" (dynamic name)');
 console.log('   2. Save Excel to server');
 console.log('   3. Load Excel data in TallyIntegrator.importToTally()');
 console.log('   4. 🎯 Apply standardization: "Input CGST @ 9%" → "CGST INPUT"');
 console.log('   5. Process standardized data for Tally');
 console.log('   6. Import uses "CGST INPUT" instead of creating new "Input CGST @ 9%" ledger');
 console.log('');
 console.log('🚀 READY FOR PRODUCTION WITH FIXED LEDGER STANDARDIZATION:');
 console.log('   ✅ Backend API: Complete with fixed ledger standardization timing');
 console.log('   ✅ Excel Generation: Creates dynamic names without standardization');
 console.log('   ✅ Tally Integration: Applies standardization at correct time during import');
 console.log('   ✅ Alias Mapping: Works with ANY aliases in master file');
 console.log('   ✅ Error Prevention: No more duplicate ledger creation');
 console.log('   ✅ File Management: Secure storage and download capabilities');
 console.log('   ✅ Analytics: Complete tracking of standardization success rates');
 console.log('   🎯 FIXED SOLUTION: Ledger standardization happens at the right time!');
 console.log('   📚 UNIVERSAL ALIASES: Any alias in master file will be properly mapped!');
});

// Disable HTTP socket timeout so long bulk-upload requests (multiple files × Claude API) never get dropped
httpServer.timeout = 0;          // no per-request timeout
httpServer.keepAliveTimeout = 0; // no keep-alive timeout

// Graceful shutdown
process.on('SIGINT', () => {
 console.log('\n🔄 Gracefully shutting down Enhanced Simplifier with Fixed Ledger Standardization...');
 
 if (db) {
     db.close((err) => {
         if (err) {
             console.error('❌ Error closing database:', err.message);
         } else {
             console.log('✅ Database connection closed');
         }
         
         console.log('👋 Enhanced Simplifier with Fixed Ledger Standardization shut down complete');
         process.exit(0);
     });
 } else {
     console.log('👋 Enhanced Simplifier with Fixed Ledger Standardization shut down complete');
     process.exit(0);
 }
});

// Error handling
process.on('uncaughtException', (error) => {
 console.error('❌ Uncaught Exception:', error);
 console.error('Stack:', error.stack);
});

process.on('unhandledRejection', (reason, promise) => {
 console.error('❌ Unhandled Rejection at:', promise, 'reason:', reason);
});

// Export for testing
module.exports = {
 app,
 db,
 DuplicateDetector,
 LineItemsValidator,
 LineItemsExtractor,
 AILineItemsProcessor,
 ParseurInvoiceProcessor,
 EnhancedAmountProcessor,
 UnifiedDocumentProcessor,
 JournalVoucherProcessor, // 🆕 New export for Payment tab
 TallyIntegrator, // 🚀 New export for Tally auto-import with fixed timing
 LedgerStandardizer, // 📚 New export for Ledger standardization
 convertToTallyFormat, // 🎯 Fixed - no standardization in this function
 prepareValue,
 prepareNumericValue,
 prepareBooleanValue
};

// --- BEGIN: Vendor Selection API ---

/**
 * API: Get all ledgers with their name, parent group, and GSTIN (for vendor selection dropdown)
 */
app.get('/api/ledgers/vendors', async (req, res) => {
    try {
        // Load the master ledger file (if not already loaded)
        await ledgerStandardizer.loadLedgerMaster();
        const ledgerMasterPath = './Ledger extractionV2/TallyData_Complete_latest.xlsx';
        const sheetName = '📋 Ledgers';
        if (!fs.existsSync(ledgerMasterPath)) {
            return res.status(404).json({ error: 'Ledger master file not found' });
        }
        const workbook = XLSX.readFile(ledgerMasterPath);
        if (!workbook.SheetNames.includes(sheetName)) {
            return res.status(404).json({ error: `Sheet '${sheetName}' not found` });
        }
        const worksheet = workbook.Sheets[sheetName];
        const ledgerData = XLSX.utils.sheet_to_json(worksheet);
        // Only return ledgers with parent group Sundry Creditors or Sundry Creditors for Expenses
        const filtered = ledgerData.filter(l => {
            const parent = (l['Parent Group'] || '').toLowerCase();
            return parent === 'sundry creditors' || parent === 'sundry creditors for expenses';
        }).map(l => ({
            name: l['Ledger Name'],
            parent: l['Parent Group'],
            gstin: l['GSTN'] || '',
        }));
        res.json(filtered);
    } catch (err) {
        console.error('Failed to load ledgers for vendor selection:', err);
        res.status(500).json({ error: 'Failed to load ledgers for vendor selection' });
    }
});
// --- END: Vendor Selection API ---

// --- BEGIN: Update selected vendor for an expense ---
app.post('/api/expenses/:id/vendor', (req, res) => {
    const expenseId = req.params.id;
    const { vendor_name } = req.body;
    if (!vendor_name) {
        return res.status(400).json({ error: 'vendor_name is required' });
    }
    const sql = 'UPDATE expenses SET vendor_name = ? WHERE id = ?';
    db.run(sql, [vendor_name, expenseId], function (err) {
        if (err) {
            console.error('Failed to update vendor for expense:', err);
            return res.status(500).json({ error: 'Failed to update vendor' });
        }
        res.json({ success: true });
    });
});
// --- END: Update selected vendor for an expense ---