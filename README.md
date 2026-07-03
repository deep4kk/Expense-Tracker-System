# Expense Tracker & Approval System

A comprehensive Google Apps Script Web Application for managing employee expense claims with multi-level approval workflow, budget tracking, and receipt management.

## Features

- **Multi-Level Approval**: Manager approval for ≤₹5,000, Manager + Finance for >₹5,000
- **Budget Middleware**: Real-time budget checking against department allocations
- **Receipt Upload**: Google Drive integration for receipt storage
- **Animated Dashboard**: Donut charts, budget utilization bars, progress steppers
- **Role-Based Access**: Employee, Manager, Finance Admin
- **Audit Trail**: Complete history of all expense state changes

## Setup

1. Create a new Google Apps Script project
2. Create sheets: Users, ExpenseClaims, DepartmentBudgets, AuditLog
3. Deploy as Web App

## Architecture

- State Machine: Draft → Pending Manager → Pending Finance → Approved → Paid
- Budget middleware validates against DepartmentBudgets sheet
- Sequential multi-level approval workflow
