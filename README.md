# Automated-Enrollment-System---CITD

Project Description: Automated Student Registration & Batched Reporting System
This project is an Automated Student Registration and Documentation System designed for the Central Institute of Tool Design (CITD) in Hyderabad. It replaces the manual registration process with a modern, efficient web application that intelligently manages document generation and communication.

Students fill out a detailed application form on a public website. Upon submission, a backend server built with Node.js and Express.js automatically performs the following tasks:

Instant Document Generation: It immediately creates two professional, filled-out PDF documents: the student's official certificate and a complete copy of their application form.

Centralized Record Keeping: It updates a master Excel spreadsheet in real-time with the new student's information, maintaining a centralized and always-current database of all registrations.

Automated & Batched Notifications: The system handles communications with both students and faculty with maximum efficiency:

An immediate confirmation email is sent to the student, with their submitted application form attached for their personal records.

For faculty, it utilizes a sophisticated scheduling system (node-cron) to queue all generated documents. Instead of sending an email for every single submission, it sends a single, consolidated report at pre-defined times each day (e.g., 12:00 PM and 9:00 PM). This batch email contains all student application forms and certificates from that period, along with the latest master Excel spreadsheet.

This system, particularly with the addition of batched reporting, streamlines the entire enrollment process. It reduces administrative workload, eliminates manual errors, prevents email clutter for faculty, and ensures all parties receive timely, consistent, and professional documentation.

