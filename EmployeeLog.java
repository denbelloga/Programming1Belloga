package com.mycompany.projectden;

import java.io.FileInputStream;
import java.io.IOException;
import java.time.Duration;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.Iterator;
import org.apache.poi.ss.usermodel.*;

public class EmployeeLog {

    public static void main(String[] args) {
        String filePath = "C:\\Users\\eyell\\OneDrive\\Documents\\NetBeansProjects\\projectDEN\\src\\main\\resources\\Copy of MotorPH Employee Data (1).xlsx";

        try (FileInputStream file = new FileInputStream(filePath); Workbook workbook = WorkbookFactory.create(file)) {
            Sheet employeeDataSheet = workbook.getSheetAt(0);
            Sheet attendanceSheet = workbook.getSheetAt(1); // Sheet 2

            Iterator<Row> employeeDataIterator = employeeDataSheet.iterator();
            employeeDataIterator.next(); // Skip header row

            System.out.printf("%-15s %-15s %-15s %-15s %-15s %-15s %-15s %-15s %-15s %-15s %-15s %-15s\n", "Employee #", "Basic Salary", "Allowance", "SSS", "Pag-IBIG", "PhilHealth", "Tax", "Total Deductions", "Overtime Pay", "Late Penalty", "Net Salary", "Hours Worked");

            while (employeeDataIterator.hasNext()) {
                Row employeeDataRow = employeeDataIterator.next();
                double employeeNumber = employeeDataRow.getCell(0).getNumericCellValue(); // Employee # (column A)
                double basicSalary = employeeDataRow.getCell(13).getNumericCellValue(); // Basic Salary (column N)

                // Read allowance from columns O, P, and Q (indices 14, 15, 16)
                double allowance1 = readNumericCellValue(employeeDataRow.getCell(14));
                double allowance2 = readNumericCellValue(employeeDataRow.getCell(15));
                double allowance3 = readNumericCellValue(employeeDataRow.getCell(16));
                double totalAllowance = allowance1 + allowance2 + allowance3;

                // Use basicSalary for deductions
                double salaryForDeductions = basicSalary;

                // Calculate deductions
                double sss = calculateSSS(salaryForDeductions);
                double pagibig = calculatePagIBIG(salaryForDeductions);
                double philhealth = calculatePhilHealth(salaryForDeductions);
                double tax = calculateWithholdingTax(salaryForDeductions);

                // Read overtime rate from column T (index 19)
                double overtimeRate = readNumericCellValue(employeeDataRow.getCell(19));

                // Get time in and out from Attendance Sheet
                LocalDateTime timeIn = getTimeFromSheet(attendanceSheet, employeeNumber, 4); // Column E (index 4)
                LocalDateTime timeOut = getTimeFromSheet(attendanceSheet, employeeNumber, 5); // Column F (index 5)
                double hoursWorked = calculateHoursWorked(timeIn, timeOut);

                double overtimePay = calculateOvertimePay(overtimeRate, hoursWorked - 8); // Overtime after 8 hours
                if (overtimePay < 0) {
                    overtimePay = 0;
                }

                double latePenalty = calculateLatePenalty(basicSalary, calculateLateMinutes(timeIn));

                // Calculate total deductions
                double totalDeductions = sss + pagibig + philhealth + tax + latePenalty;

                double netSalary = (basicSalary + totalAllowance + overtimePay) - totalDeductions;

                System.out.printf("%-15.0f %-15.2f %-15.2f %-15.2f %-15.2f %-15.2f %-15.2f %-15.2f %-15.2f %-15.2f %-15.2f %-15.2f\n", employeeNumber, basicSalary, totalAllowance, sss, pagibig, philhealth, tax, totalDeductions, overtimePay, latePenalty, netSalary, hoursWorked);
            }
        } catch (IOException e) {
            System.out.println("Error reading the Excel file: " + e.getMessage());
        }
    }

    private static LocalDateTime getTimeFromSheet(Sheet sheet, double employeeNumber, int columnIndex) {
        Iterator<Row> rowIterator = sheet.iterator();
        rowIterator.next(); // Skip header row
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            if (row.getCell(0).getNumericCellValue() == employeeNumber) {
                Cell cell = row.getCell(columnIndex);
                if (cell.getCellType() == CellType.NUMERIC) {
                    return cell.getLocalDateTimeCellValue();
                } else if (cell.getCellType() == CellType.STRING) {
                    try {
                        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss");
                        return LocalDateTime.parse(cell.getStringCellValue(), formatter);
                    } catch (Exception e) {
                        System.err.println("Error parsing date/time: " + cell.getStringCellValue());
                        return null;
                    }
                }
            }
        }
        return null;
    }

    private static double calculateHoursWorked(LocalDateTime timeIn, LocalDateTime timeOut) {
        if (timeIn == null || timeOut == null) {
            return 0;
        }
        Duration duration = Duration.between(timeIn, timeOut);
        double hours = duration.toMinutes() / 60.0;
        return Math.max(0, hours - 1); // Subtract 1 hour for lunch break
    }

    private static double calculateLateMinutes(LocalDateTime timeIn) {
        if (timeIn == null) {
            return 0;
        }
        LocalDateTime standardTimeIn = timeIn.toLocalDate().atTime(8, 30);
        if (timeIn.isAfter(standardTimeIn)) {
            Duration duration = Duration.between(standardTimeIn, timeIn);
            return duration.toMinutes();
        }
        return 0;
    }

    private static double readNumericCellValue(Cell cell) {
        if (cell == null) {
            return 0; // Handle null cells
        }
        if (cell.getCellType() == CellType.NUMERIC) {
            return cell.getNumericCellValue();
        } else if (cell.getCellType() == CellType.STRING) {
            try {
                return Double.parseDouble(cell.getStringCellValue());
            } catch (NumberFormatException e) {
                System.err.println("Error: Cell is not a valid number: " + cell.getStringCellValue());
                return 0; // Default to 0 or handle the error as needed
            }
        } else {
            return 0; // Default to 0 for other cell types
        }
    }

    public static double calculateOvertimePay(double overtimeRate, double hoursOvertime) {
        return hoursOvertime * overtimeRate;
    }

    public static double calculateLatePenalty(double salary, double minutesLate) {
        if (minutesLate <= 10) return 0;
        double hourlyRate = salary / 160;
        double minuteRate = hourlyRate / 60;
        return (minutesLate - 10) * minuteRate;
    }

    public static double calculateSSS(double salary) {
        return Math.min(1350.00, salary * 0.045);
    }

    public static double calculatePagIBIG(double salary) {
        return Math.min(100.00, salary * 0.02);
    }

    public static double calculatePhilHealth(double salary) {
        return Math.min(900.00, salary * 0.04 / 2);
    }

    public static double calculateWithholdingTax(double salary) {
        if (salary <= 20833) return 0;
        else if (salary <= 33333) return (salary - 20833) * 0.20;
        else if (salary <= 66667) return 2500 + (salary - 33333) * 0.25;
        else if (salary <= 166667) return 10833 + (salary - 66667) * 0.30;
        else if (salary <= 666667) return 40833 + (salary - 166667) * 0.32;
        else return 200833 + (salary - 666667) * 0.35;
    }
}
