<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Business Management System - Task Manager, Payroll, Finance, POS & AI Chat</title>
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
    <style>
        :root {
            --primary: #2563eb;
            --secondary: #1d4ed8;
            --success: #10b981;
            --danger: #ef4444;
            --warning: #f59e0b;
            --info: #3b82f6;
            --light: #f8fafc;
            --dark: #1e293b;
            --gray: #64748b;
            --white: #ffffff;
            --card-bg: #ffffff;
            --body-bg: #f1f5f9;
            --text-color: #334155;
            --border-color: #e2e8f0;
            --shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
            --radius: 10px;
        }

        .dark-mode {
            --primary: #3b82f6;
            --secondary: #2563eb;
            --success: #34d399;
            --danger: #f87171;
            --warning: #fbbf24;
            --info: #60a5fa;
            --light: #334155;
            --dark: #f8fafc;
            --gray: #94a3b8;
            --white: #1e2937;
            --card-bg: #334155;
            --body-bg: #0f172a;
            --text-color: #f8fafc;
            --border-color: #475569;
            --shadow: 0 4px 6px rgba(0, 0, 0, 0.3);
        }

        * {
            box-sizing: border-box;
            margin: 0;
            padding: 0;
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            transition: background-color 0.3s, color 0.3s;
        }

        body {
            background-color: var(--body-bg);
            color: var(--text-color);
            line-height: 1.6;
        }

        .container {
            max-width: 1400px;
            margin: 0 auto;
            padding: 20px;
        }

        header {
            text-align: center;
            margin-bottom: 30px;
            padding: 25px 0;
            background: linear-gradient(135deg, var(--primary), var(--secondary));
            color: white;
            border-radius: var(--radius);
            box-shadow: var(--shadow);
            position: relative;
        }

        h1 {
            font-size: 2.5rem;
            margin-bottom: 10px;
        }

        .theme-toggle {
            position: absolute;
            top: 20px;
            right: 20px;
            background: rgba(255, 255, 255, 0.2);
            border: none;
            color: white;
            padding: 10px;
            border-radius: 50%;
            cursor: pointer;
            font-size: 1.2rem;
        }

        .dashboard {
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: center;
            min-height: 70vh;
        }

        .module-buttons {
            display: flex;
            flex-wrap: wrap;
            justify-content: center;
            gap: 30px;
            margin-top: 30px;
        }

        .module-btn {
            background-color: var(--primary);
            color: white;
            border: none;
            padding: 40px 50px;
            border-radius: var(--radius);
            cursor: pointer;
            font-size: 1.8rem;
            font-weight: bold;
            box-shadow: var(--shadow);
            transition: all 0.3s ease;
            display: flex;
            flex-direction: column;
            align-items: center;
            gap: 15px;
            width: 300px;
            height: 300px;
            justify-content: center;
        }

        .module-btn:hover {
            background-color: var(--secondary);
            transform: translateY(-5px);
        }

        .module-btn i {
            font-size: 4rem;
        }

        .module-content {
            display: none;
            background-color: var(--card-bg);
            padding: 25px;
            border-radius: var(--radius);
            box-shadow: var(--shadow);
            margin-bottom: 20px;
        }

        .module-content.active {
            display: block;
        }

        .back-btn {
            background-color: var(--gray);
            color: white;
            border: none;
            padding: 10px 20px;
            border-radius: 5px;
            cursor: pointer;
            font-size: 1rem;
            margin-bottom: 20px;
        }

        .back-btn:hover {
            background-color: #5a6268;
        }

        .tabs {
            display: flex;
            margin-bottom: 20px;
            background-color: var(--card-bg);
            border-radius: var(--radius);
            box-shadow: var(--shadow);
            overflow: hidden;
        }

        .tab {
            flex: 1;
            padding: 15px;
            text-align: center;
            cursor: pointer;
            transition: background-color 0.3s;
            border-bottom: 3px solid transparent;
        }

        .tab:hover {
            background-color: var(--light);
        }

        .tab.active {
            border-bottom: 3px solid var(--primary);
            background-color: var(--light);
        }

        .tab-content {
            display: none;
            background-color: var(--card-bg);
            padding: 25px;
            border-radius: var(--radius);
            box-shadow: var(--shadow);
            margin-bottom: 20px;
        }

        .tab-content.active {
            display: block;
        }

        .form-group {
            margin-bottom: 15px;
        }

        label {
            display: block;
            margin-bottom: 5px;
            font-weight: 600;
            color: var(--text-color);
        }

        input, textarea, select {
            width: 100%;
            padding: 12px;
            border: 1px solid var(--border-color);
            border-radius: 5px;
            font-size: 1rem;
            background-color: var(--card-bg);
            color: var(--text-color);
        }

        textarea {
            resize: vertical;
            min-height: 100px;
        }

        button {
            background-color: var(--primary);
            color: white;
            border: none;
            padding: 12px 20px;
            border-radius: 5px;
            cursor: pointer;
            font-size: 1rem;
            transition: background-color 0.3s;
            margin-right: 10px;
            margin-bottom: 10px;
        }

        button:hover {
            background-color: var(--secondary);
        }

        .btn-secondary {
            background-color: var(--gray);
        }

        .btn-secondary:hover {
            background-color: #5a6268;
        }

        .btn-danger {
            background-color: var(--danger);
        }

        .btn-danger:hover {
            background-color: #c53030;
        }

        .btn-success {
            background-color: var(--success);
        }

        .btn-success:hover {
            background-color: #059669;
        }

        .btn-warning {
            background-color: var(--warning);
        }

        .btn-warning:hover {
            background-color: #d97706;
        }

        .data-table {
            width: 100%;
            border-collapse: collapse;
            margin-top: 20px;
        }

        .data-table th, .data-table td {
            padding: 12px;
            text-align: left;
            border-bottom: 1px solid var(--border-color);
        }

        .data-table th {
            background-color: var(--light);
            font-weight: 600;
        }

        .data-table tr:hover {
            background-color: var(--light);
        }

        .notification {
            position: fixed;
            top: 20px;
            right: 20px;
            padding: 15px 20px;
            background-color: var(--success);
            color: white;
            border-radius: 5px;
            box-shadow: var(--shadow);
            transform: translateX(150%);
            transition: transform 0.3s ease-out;
            z-index: 1000;
        }

        .notification.show {
            transform: translateX(0);
        }

        .notification.error {
            background-color: var(--danger);
        }

        .modal {
            display: none;
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background-color: rgba(0, 0, 0, 0.5);
            z-index: 1000;
            overflow: auto;
        }

        .modal-content {
            background-color: var(--card-bg);
            margin: 5% auto;
            padding: 20px;
            border-radius: var(--radius);
            width: 90%;
            max-width: 800px;
            position: relative;
        }

        .close-modal {
            position: absolute;
            top: 10px;
            right: 15px;
            font-size: 24px;
            cursor: pointer;
            color: var(--text-color);
        }

        .card {
            background-color: var(--card-bg);
            border-radius: var(--radius);
            box-shadow: var(--shadow);
            padding: 20px;
            margin-bottom: 20px;
        }

        .card-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 15px;
            border-bottom: 1px solid var(--border-color);
            padding-bottom: 10px;
        }

        .card-title {
            font-size: 1.3rem;
            font-weight: 600;
            color: var(--text-color);
        }

        .card-body {
            color: var(--text-color);
        }

        .summary-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
            gap: 20px;
            margin-bottom: 20px;
        }

        .summary-card {
            background-color: var(--card-bg);
            border-radius: var(--radius);
            box-shadow: var(--shadow);
            padding: 20px;
            text-align: center;
        }

        .summary-card h3 {
            font-size: 1rem;
            color: var(--gray);
            margin-bottom: 10px;
        }

        .summary-card p {
            font-size: 1.8rem;
            font-weight: bold;
            color: var(--primary);
        }

        .chart-container {
            height: 300px;
            margin-top: 20px;
        }

        .product-search {
            position: relative;
            margin-bottom: 20px;
        }

        .search-results {
            position: absolute;
            top: 100%;
            left: 0;
            right: 0;
            background-color: var(--card-bg);
            border: 1px solid var(--border-color);
            border-radius: 5px;
            max-height: 200px;
            overflow-y: auto;
            z-index: 100;
            box-shadow: var(--shadow);
        }

        .search-result-item {
            padding: 10px;
            cursor: pointer;
            border-bottom: 1px solid var(--border-color);
        }

        .search-result-item:hover {
            background-color: var(--light);
        }

        .search-result-item:last-child {
            border-bottom: none;
        }

        .selected-products {
            margin-top: 20px;
        }

        .product-row {
            display: grid;
            grid-template-columns: 2fr 1fr 1fr 1fr auto;
            gap: 10px;
            margin-bottom: 10px;
            align-items: center;
        }

        .product-row input {
            padding: 8px;
        }

        .sale-summary {
            margin-top: 20px;
            padding: 15px;
            background-color: var(--light);
            border-radius: 5px;
        }

        .sale-summary-row {
            display: flex;
            justify-content: space-between;
            margin-bottom: 10px;
        }

        .sale-summary-row.total {
            font-weight: bold;
            font-size: 1.2rem;
            border-top: 1px solid var(--border-color);
            padding-top: 10px;
        }

        /* Task Manager Styles */
        .task-container {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 20px;
        }

        .task-section {
            background-color: var(--card-bg);
            border-radius: var(--radius);
            box-shadow: var(--shadow);
            padding: 20px;
        }

        .task-section-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 15px;
            border-bottom: 1px solid var(--border-color);
            padding-bottom: 10px;
        }

        .task-list {
            max-height: 500px;
            overflow-y: auto;
        }

        .task-item {
            background-color: var(--light);
            border-radius: 8px;
            padding: 15px;
            margin-bottom: 10px;
            position: relative;
            cursor: pointer;
        }

        .task-item:hover {
            transform: translateY(-2px);
            box-shadow: var(--shadow);
        }

        .task-item h4 {
            margin-bottom: 10px;
            color: var(--text-color);
        }

        .task-meta {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 10px;
            margin-bottom: 10px;
        }

        .task-meta div {
            display: flex;
            flex-direction: column;
        }

        .task-meta label {
            font-size: 0.8rem;
            color: var(--gray);
            margin-bottom: 2px;
        }

        .task-meta span {
            font-weight: 600;
        }

        .task-actions {
            display: flex;
            justify-content: flex-end;
            gap: 10px;
        }

        .task-status {
            position: absolute;
            top: 10px;
            right: 10px;
            padding: 3px 8px;
            border-radius: 12px;
            font-size: 0.75rem;
            font-weight: bold;
        }

        .status-pending {
            background-color: var(--warning);
            color: white;
        }

        .status-completed {
            background-color: var(--success);
            color: white;
        }

        .task-form {
            background-color: var(--card-bg);
            border-radius: var(--radius);
            box-shadow: var(--shadow);
            padding: 20px;
            margin-bottom: 20px;
        }

        .task-form-grid {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 15px;
        }

        .task-detail-section {
            margin-bottom: 20px;
        }

        .task-detail-section h4 {
            margin-bottom: 10px;
            color: var(--text-color);
            border-bottom: 1px solid var(--border-color);
            padding-bottom: 5px;
        }

        .comment-item {
            background-color: var(--light);
            border-radius: 8px;
            padding: 10px;
            margin-bottom: 10px;
        }

        .comment-meta {
            font-size: 0.8rem;
            color: var(--gray);
            margin-bottom: 5px;
        }

        .attachment-item {
            display: flex;
            align-items: center;
            background-color: var(--light);
            border-radius: 8px;
            padding: 10px;
            margin-bottom: 10px;
        }

        .attachment-item i {
            margin-right: 10px;
            font-size: 1.2rem;
            color: var(--primary);
        }

        .attachment-name {
            flex-grow: 1;
        }

        .attachment-actions {
            display: flex;
            gap: 5px;
        }

        .attachment-actions button {
            padding: 5px 10px;
            font-size: 0.8rem;
            margin: 0;
        }

        .file-upload {
            position: relative;
            display: inline-block;
            cursor: pointer;
            overflow: hidden;
            background-color: var(--primary);
            color: white;
            padding: 10px 15px;
            border-radius: 5px;
            margin-right: 10px;
        }

        .file-upload input[type=file] {
            position: absolute;
            left: 0;
            top: 0;
            opacity: 0;
            width: 100%;
            height: 100%;
            cursor: pointer;
        }

        .file-upload:hover {
            background-color: var(--secondary);
        }

        .invoice-actions {
            display: flex;
            justify-content: center;
            gap: 15px;
            margin-top: 20px;
        }

        .invoice-preview {
            background-color: var(--light);
            border-radius: 8px;
            padding: 20px;
            margin-top: 20px;
        }

        .invoice-header {
            text-align: center;
            margin-bottom: 20px;
            border-bottom: 1px solid var(--border-color);
            padding-bottom: 10px;
        }

        .invoice-title {
            font-size: 1.5rem;
            font-weight: bold;
            margin-bottom: 5px;
        }

        .invoice-meta {
            display: flex;
            justify-content: space-between;
            margin-bottom: 20px;
        }

        .invoice-details {
            margin-bottom: 20px;
        }

        .invoice-row {
            display: flex;
            justify-content: space-between;
            margin-bottom: 10px;
        }

        .invoice-row.total {
            font-weight: bold;
            font-size: 1.2rem;
            border-top: 1px solid var(--border-color);
            padding-top: 10px;
        }

        .task-overdue {
            border-left: 4px solid var(--danger);
        }

        .task-due-soon {
            border-left: 4px solid var(--warning);
        }

        .task-normal {
            border-left: 4px solid var(--success);
        }

        /* AI Chat Styles */
        .chat-container {
            display: flex;
            flex-direction: column;
            height: 70vh;
            background-color: var(--card-bg);
            border-radius: var(--radius);
            box-shadow: var(--shadow);
            overflow: hidden;
        }

        .chat-header {
            padding: 15px 20px;
            background-color: var(--primary);
            color: white;
            display: flex;
            align-items: center;
            gap: 10px;
        }

        .chat-messages {
            flex-grow: 1;
            overflow-y: auto;
            padding: 20px;
            display: flex;
            flex-direction: column;
            gap: 15px;
        }

        .message {
            max-width: 80%;
            padding: 12px 16px;
            border-radius: 18px;
            position: relative;
        }

        .user-message {
            align-self: flex-end;
            background-color: var(--primary);
            color: white;
        }

        .ai-message {
            align-self: flex-start;
            background-color: var(--light);
            color: var(--text-color);
        }

        .message-content {
            margin-bottom: 10px;
        }

        .message-actions {
            display: flex;
            gap: 8px;
            margin-top: 8px;
        }

        .message-actions button {
            font-size: 0.8rem;
            padding: 5px 10px;
            margin: 0;
        }

        .chat-input-container {
            display: flex;
            padding: 15px;
            border-top: 1px solid var(--border-color);
        }

        .chat-input {
            flex-grow: 1;
            padding: 12px 16px;
            border: 1px solid var(--border-color);
            border-radius: 24px;
            margin-right: 10px;
        }

        .send-button {
            background-color: var(--primary);
            color: white;
            border: none;
            border-radius: 50%;
            width: 48px;
            height: 48px;
            display: flex;
            align-items: center;
            justify-content: center;
            cursor: pointer;
        }

        .send-button:hover {
            background-color: var(--secondary);
        }

        .chat-response-chart {
            margin-top: 15px;
            height: 250px;
        }

        .chat-response-table {
            margin-top: 15px;
            width: 100%;
            border-collapse: collapse;
        }

        .chat-response-table th, .chat-response-table td {
            padding: 8px 12px;
            text-align: left;
            border-bottom: 1px solid var(--border-color);
        }

        .chat-response-table th {
            background-color: var(--light);
            font-weight: 600;
        }

        .typing-indicator {
            display: inline-flex;
            align-items: center;
            padding: 8px 12px;
            background-color: var(--light);
            border-radius: 18px;
        }

        .typing-indicator span {
            height: 8px;
            width: 8px;
            border-radius: 50%;
            background-color: var(--gray);
            display: inline-block;
            margin: 0 2px;
            animation: typing 1.4s infinite;
        }

        .typing-indicator span:nth-child(2) {
            animation-delay: 0.2s;
        }

        .typing-indicator span:nth-child(3) {
            animation-delay: 0.4s;
        }

        @keyframes typing {
            0%, 60%, 100% {
                transform: translateY(0);
            }
            30% {
                transform: translateY(-10px);
            }
        }

        @media (max-width: 768px) {
            .module-buttons {
                flex-direction: column;
                align-items: center;
            }
            
            .module-btn {
                width: 100%;
                max-width: 300px;
            }
            
            .tabs {
                flex-direction: column;
            }
            
            .modal-content {
                width: 95%;
                margin: 5% auto;
            }
            
            .product-row {
                grid-template-columns: 1fr;
                gap: 5px;
            }
            
            .task-container {
                grid-template-columns: 1fr;
            }
            
            .task-form-grid {
                grid-template-columns: 1fr;
            }
            
            .invoice-meta {
                flex-direction: column;
            }
            
            .chat-container {
                height: 80vh;
            }
        }
    </style>
</head>
<body>
    <div class="container">
        <header>
            <h1><i class="fas fa-briefcase"></i> Business Management System</h1>
            <p>Task Management, Payroll, Finance, POS & AI Chat</p>
            <button class="theme-toggle" id="theme-toggle">
                <i class="fas fa-moon"></i>
            </button>
        </header>

        <div id="dashboard" class="dashboard">
            <h2>Select a Module</h2>
            <div class="module-buttons">
                <button class="module-btn" id="task-module-btn">
                    <i class="fas fa-tasks"></i>
                    Task Manager
                </button>
                <button class="module-btn" id="payroll-module-btn">
                    <i class="fas fa-money-check-alt"></i>
                    Payroll
                </button>
                <button class="module-btn" id="financial-module-btn">
                    <i class="fas fa-chart-line"></i>
                    Business Financial
                </button>
                <button class="module-btn" id="pos-module-btn">
                    <i class="fas fa-cash-register"></i>
                    POS
                </button>
                <button class="module-btn" id="ai-chat-module-btn">
                    <i class="fas fa-robot"></i>
                    AI Chat
                </button>
            </div>
        </div>

        <!-- Task Manager Module -->
        <div id="task-module" class="module-content">
            <button class="back-btn" id="task-back-btn"><i class="fas fa-arrow-left"></i> Back to Dashboard</button>
            <h2>Task Manager</h2>
            
            <div class="task-form">
                <h3>Add New Task</h3>
                <form id="task-form">
                    <div class="task-form-grid">
                        <div class="form-group">
                            <label for="task-name">Task Name</label>
                            <input type="text" id="task-name" required>
                        </div>
                        
                        <div class="form-group">
                            <label for="task-category">Category</label>
                            <select id="task-category" required>
                                <option value="">Select Category</option>
                                <option value="Important">Important</option>
                                <option value="Very Important">Very Important</option>
                                <option value="Critical">Critical</option>
                                <option value="Normal">Normal</option>
                                <option value="Low">Low</option>
                                <option value="Other">Other</option>
                            </select>
                        </div>
                        
                        <div class="form-group">
                            <label for="task-type">Type</label>
                            <select id="task-type" required>
                                <option value="">Select Type</option>
                                <option value="Fix">Fix</option>
                                <option value="Bugs">Bugs</option>
                                <option value="Development">Development</option>
                                <option value="Changes">Changes</option>
                                <option value="Design">Design</option>
                                <option value="Add">Add</option>
                                <option value="Make">Make</option>
                            </select>
                        </div>
                        
                        <div class="form-group">
                            <label for="task-cost">Cost (Rs.)</label>
                            <input type="number" id="task-cost" min="0" step="0.01" required>
                        </div>
                    </div>
                    
                    <div class="form-group">
                        <label for="task-due-date">Due Date</label>
                        <input type="date" id="task-due-date" required>
                    </div>
                    
                    <div class="form-group">
                        <label for="task-comment">Initial Comment</label>
                        <textarea id="task-comment" placeholder="Add initial comment..."></textarea>
                    </div>
                    
                    <div class="form-group">
                        <label>Attachments</label>
                        <div class="file-upload">
                            <i class="fas fa-paperclip"></i> Attach Files
                            <input type="file" id="task-attachments" multiple>
                        </div>
                        <div id="attachment-list"></div>
                    </div>
                    
                    <button type="submit">Add Task</button>
                </form>
            </div>
            
            <div class="task-container">
                <div class="task-section">
                    <div class="task-section-header">
                        <h3>Pending Tasks</h3>
                        <span class="section-count" id="pending-count">0</span>
                    </div>
                    <div id="pending-tasks" class="task-list">
                        <!-- Pending tasks will be inserted here -->
                    </div>
                </div>
                
                <div class="task-section">
                    <div class="task-section-header">
                        <h3>Completed Tasks</h3>
                        <span class="section-count" id="completed-count">0</span>
                    </div>
                    <div id="completed-tasks" class="task-list">
                        <!-- Completed tasks will be inserted here -->
                    </div>
                </div>
            </div>
        </div>

        <!-- AI Chat Module -->
        <div id="ai-chat-module" class="module-content">
            <button class="back-btn" id="ai-chat-back-btn"><i class="fas fa-arrow-left"></i> Back to Dashboard</button>
            <h2>AI Chat Assistant</h2>
            
            <div class="chat-container">
                <div class="chat-header">
                    <i class="fas fa-robot"></i>
                    <h3>Business Assistant</h3>
                </div>
                
                <div id="chat-messages" class="chat-messages">
                    <div class="message ai-message">
                        <div class="message-content">
                            Hello! I'm your Business Assistant. You can ask me questions about your business data, such as sales summaries, expenses, tasks, and more. Try asking things like "Show me today's sales summary" or "Generate a chart of expenses by category".
                        </div>
                    </div>
                </div>
                
                <div class="chat-input-container">
                    <input type="text" id="chat-input" class="chat-input" placeholder="Ask a question about your business data...">
                    <button id="send-chat-btn" class="send-button">
                        <i class="fas fa-paper-plane"></i>
                    </button>
                </div>
            </div>
        </div>

        <!-- Payroll Module -->
        <div id="payroll-module" class="module-content">
            <button class="back-btn" id="payroll-back-btn"><i class="fas fa-arrow-left"></i> Back to Dashboard</button>
            <h2>Payroll Management</h2>
            
            <div class="tabs">
                <div class="tab active" data-tab="employees">Employees</div>
                <div class="tab" data-tab="payslips">Payslips</div>
                <div class="tab" data-tab="reports">Reports</div>
            </div>
            
            <div id="employees-tab" class="tab-content active">
                <div style="display: flex; justify-content: space-between; margin-bottom: 20px;">
                    <h3>Employee List</h3>
                    <button id="add-employee-btn"><i class="fas fa-plus"></i> Add Employee</button>
                </div>
                
                <table class="data-table">
                    <thead>
                        <tr>
                            <th>ID</th>
                            <th>Name</th>
                            <th>Contact</th>
                            <th>Salary (Rs.)</th>
                            <th>Join Date</th>
                            <th>Actions</th>
                        </tr>
                    </thead>
                    <tbody id="employees-table-body">
                        <!-- Employee rows will be inserted here -->
                    </tbody>
                </table>
            </div>
            
            <div id="payslips-tab" class="tab-content">
                <div style="display: flex; justify-content: space-between; margin-bottom: 20px;">
                    <h3>Generate Payslip</h3>
                </div>
                
                <form id="payslip-form">
                    <div class="form-group">
                        <label for="employee-select">Employee</label>
                        <select id="employee-select" required>
                            <option value="">Select Employee</option>
                        </select>
                    </div>
                    
                    <div class="form-group">
                        <label for="month">Month</label>
                        <select id="month" required>
                            <option value="January">January</option>
                            <option value="February">February</option>
                            <option value="March">March</option>
                            <option value="April">April</option>
                            <option value="May">May</option>
                            <option value="June">June</option>
                            <option value="July">July</option>
                            <option value="August">August</option>
                            <option value="September">September</option>
                            <option value="October">October</option>
                            <option value="November">November</option>
                            <option value="December">December</option>
                        </select>
                    </div>
                    
                    <div class="form-group">
                        <label for="year">Year</label>
                        <input type="number" id="year" min="2020" max="2030" value="2023" required>
                    </div>
                    
                    <div class="form-group">
                        <label for="overtime-hours">Overtime Hours</label>
                        <input type="number" id="overtime-hours" min="0" step="0.5" value="0">
                    </div>
                    
                    <div class="form-group">
                        <label for="deductions">Deductions (Rs.)</label>
                        <input type="number" id="deductions" min="0" step="0.01" value="0">
                    </div>
                    
                    <button type="submit">Generate Payslip</button>
                </form>
            </div>
            
            <div id="reports-tab" class="tab-content">
                <div style="display: flex; justify-content: space-between; margin-bottom: 20px;">
                    <h3>Payroll Reports</h3>
                    <div>
                        <button id="payroll-summary-btn">Summary Report</button>
                        <button id="export-payroll-btn">Export to Excel</button>
                    </div>
                </div>
                
                <div class="card">
                    <div class="card-header">
                        <h3 class="card-title">Monthly Salary Report</h3>
                    </div>
                    <div class="card-body">
                        <div class="form-group">
                            <label for="report-month">Month</label>
                            <select id="report-month">
                                <option value="January">January</option>
                                <option value="February">February</option>
                                <option value="March">March</option>
                                <option value="April">April</option>
                                <option value="May">May</option>
                                <option value="June">June</option>
                                <option value="July">July</option>
                                <option value="August">August</option>
                                <option value="September">September</option>
                                <option value="October">October</option>
                                <option value="November">November</option>
                                <option value="December">December</option>
                            </select>
                        </div>
                        
                        <div class="form-group">
                            <label for="report-year">Year</label>
                            <input type="number" id="report-year" min="2020" max="2030" value="2023">
                        </div>
                        
                        <button id="generate-payroll-report-btn">Generate Report</button>
                        
                        <div id="payroll-report-container" style="margin-top: 20px;">
                            <!-- Report will be displayed here -->
                        </div>
                    </div>
                </div>
            </div>
        </div>

        <!-- Business Financial Module -->
        <div id="financial-module" class="module-content">
            <button class="back-btn" id="financial-back-btn"><i class="fas fa-arrow-left"></i> Back to Dashboard</button>
            <h2>Business Financial Management</h2>
            
            <div class="tabs">
                <div class="tab active" data-tab="income">Income</div>
                <div class="tab" data-tab="expense">Expenses</div>
                <div class="tab" data-tab="summary">Summary</div>
                <div class="tab" data-tab="reports">Reports</div>
            </div>
            
            <div id="income-tab" class="tab-content active">
                <div style="display: flex; justify-content: space-between; margin-bottom: 20px;">
                    <h3>Income Records</h3>
                    <button id="add-income-btn"><i class="fas fa-plus"></i> Add Income</button>
                </div>
                
                <table class="data-table">
                    <thead>
                        <tr>
                            <th>Source</th>
                            <th>Amount (Rs.)</th>
                            <th>Date</th>
                            <th>Description</th>
                            <th>Actions</th>
                        </tr>
                    </thead>
                    <tbody id="income-table-body">
                        <!-- Income rows will be inserted here -->
                    </tbody>
                </table>
            </div>
            
            <div id="expense-tab" class="tab-content">
                <div style="display: flex; justify-content: space-between; margin-bottom: 20px;">
                    <h3>Expense Records</h3>
                    <button id="add-expense-btn"><i class="fas fa-plus"></i> Add Expense</button>
                </div>
                
                <table class="data-table">
                    <thead>
                        <tr>
                            <th>Category</th>
                            <th>Amount (Rs.)</th>
                            <th>Date</th>
                            <th>Description</th>
                            <th>Actions</th>
                        </tr>
                    </thead>
                    <tbody id="expense-table-body">
                        <!-- Expense rows will be inserted here -->
                    </tbody>
                </table>
            </div>
            
            <div id="summary-tab" class="tab-content">
                <h3>Financial Summary</h3>
                
                <div style="display: flex; justify-content: space-between; margin-bottom: 20px;">
                    <div>
                        <label for="summary-period">Period</label>
                        <select id="summary-period">
                            <option value="today">Today</option>
                            <option value="month">This Month</option>
                            <option value="year">This Year</option>
                        </select>
                    </div>
                    <button id="refresh-summary-btn"><i class="fas fa-sync-alt"></i> Refresh</button>
                </div>
                
                <div class="summary-grid">
                    <div class="summary-card">
                        <h3>Total Income</h3>
                        <p id="total-income">Rs. 0.00</p>
                    </div>
                    <div class="summary-card">
                        <h3>Total Expenses</h3>
                        <p id="total-expenses">Rs. 0.00</p>
                    </div>
                    <div class="summary-card">
                        <h3>Profit/Loss</h3>
                        <p id="profit-loss">Rs. 0.00</p>
                    </div>
                </div>
                
                <div class="chart-container">
                    <canvas id="financial-chart"></canvas>
                </div>
            </div>
            
            <div id="reports-tab" class="tab-content">
                <div style="display: flex; justify-content: space-between; margin-bottom: 20px;">
                    <h3>Financial Reports</h3>
                    <div>
                        <button id="daily-report-btn">Daily Report</button>
                        <button id="monthly-report-btn">Monthly Report</button>
                        <button id="export-finance-btn">Export to Excel</button>
                    </div>
                </div>
                
                <div class="card">
                    <div class="card-header">
                        <h3 class="card-title">Report Options</h3>
                    </div>
                    <div class="card-body">
                        <div class="form-group">
                            <label for="report-type">Report Type</label>
                            <select id="report-type">
                                <option value="income">Income Report</option>
                                <option value="expense">Expense Report</option>
                                <option value="profit-loss">Profit/Loss Report</option>
                            </select>
                        </div>
                        
                        <div class="form-group">
                            <label for="report-date-range">Date Range</label>
                            <select id="report-date-range">
                                <option value="today">Today</option>
                                <option value="week">This Week</option>
                                <option value="month">This Month</option>
                                <option value="year">This Year</option>
                                <option value="custom">Custom Range</option>
                            </select>
                        </div>
                        
                        <div id="custom-date-range" style="display: none;">
                            <div class="form-group">
                                <label for="start-date">Start Date</label>
                                <input type="date" id="start-date">
                            </div>
                            
                            <div class="form-group">
                                <label for="end-date">End Date</label>
                                <input type="date" id="end-date">
                            </div>
                        </div>
                        
                        <button id="generate-finance-report-btn">Generate Report</button>
                        
                        <div id="finance-report-container" style="margin-top: 20px;">
                            <!-- Report will be displayed here -->
                        </div>
                    </div>
                </div>
            </div>
        </div>

        <!-- POS Module -->
        <div id="pos-module" class="module-content">
            <button class="back-btn" id="pos-back-btn"><i class="fas fa-arrow-left"></i> Back to Dashboard</button>
            <h2>Point of Sale</h2>
            
            <div class="tabs">
                <div class="tab active" data-tab="products">Products</div>
                <div class="tab" data-tab="sales">Sales</div>
                <div class="tab" data-tab="reports">Reports</div>
            </div>
            
            <div id="products-tab" class="tab-content active">
                <div style="display: flex; justify-content: space-between; margin-bottom: 20px;">
                    <h3>Product List</h3>
                    <button id="add-product-btn"><i class="fas fa-plus"></i> Add Product</button>
                </div>
                
                <table class="data-table">
                    <thead>
                        <tr>
                            <th>ID</th>
                            <th>Name</th>
                            <th>Brand</th>
                            <th>Category</th>
                            <th>Price (Rs.)</th>
                            <th>Stock</th>
                            <th>Expiry Date</th>
                            <th>Actions</th>
                        </tr>
                    </thead>
                    <tbody id="products-table-body">
                        <!-- Product rows will be inserted here -->
                    </tbody>
                </table>
            </div>
            
            <div id="sales-tab" class="tab-content">
                <div style="display: flex; justify-content: space-between; margin-bottom: 20px;">
                    <h3>New Sale</h3>
                </div>
                
                <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 20px;">
                    <div>
                        <h4>Add Products</h4>
                        <div class="product-search">
                            <input type="text" id="product-search" placeholder="Search products by name, brand, or barcode...">
                            <div id="product-search-results" class="search-results"></div>
                        </div>
                        
                        <div id="selected-products" class="selected-products">
                            <h4>Selected Products</h4>
                            <div id="selected-products-list">
                                <!-- Selected products will be inserted here -->
                            </div>
                        </div>
                    </div>
                    
                    <div>
                        <h4>Sale Details</h4>
                        <form id="sale-form">
                            <div class="form-group">
                                <label for="discount">Discount (Rs.)</label>
                                <input type="number" id="discount" min="0" step="0.01" value="0">
                            </div>
                            
                            <div class="form-group">
                                <label for="payment-method">Payment Method</label>
                                <select id="payment-method" required>
                                    <option value="Cash">Cash</option>
                                    <option value="Credit Card">Credit Card</option>
                                    <option value="Debit Card">Debit Card</option>
                                    <option value="Mobile Payment">Mobile Payment</option>
                                </select>
                            </div>
                            
                            <div class="sale-summary">
                                <div class="sale-summary-row">
                                    <span>Subtotal:</span>
                                    <span id="subtotal">Rs. 0.00</span>
                                </div>
                                <div class="sale-summary-row">
                                    <span>Discount:</span>
                                    <span id="discount-amount">Rs. 0.00</span>
                                </div>
                                <div class="sale-summary-row total">
                                    <span>Total:</span>
                                    <span id="total-amount">Rs. 0.00</span>
                                </div>
                            </div>
                            
                            <button type="submit">Complete Sale</button>
                        </form>
                    </div>
                </div>
            </div>
            
            <div id="reports-tab" class="tab-content">
                <div style="display: flex; justify-content: space-between; margin-bottom: 20px;">
                    <h3>Sales Reports</h3>
                    <div>
                        <button id="today-sales-btn">Today's Sales</button>
                        <button id="monthly-sales-btn">Monthly Sales</button>
                        <button id="low-stock-btn">Low Stock</button>
                        <button id="expiry-report-btn">Expiry Report</button>
                        <button id="export-sales-btn">Export to Excel</button>
                    </div>
                </div>
                
                <div class="card">
                    <div class="card-header">
                        <h3 class="card-title">Report Options</h3>
                    </div>
                    <div class="card-body">
                        <div class="form-group">
                            <label for="sales-report-type">Report Type</label>
                            <select id="sales-report-type">
                                <option value="today">Today's Sales</option>
                                <option value="stock-added">Today's Stock Added</option>
                                <option value="monthly-sales">Monthly Sales</option>
                                <option value="monthly-stock">Monthly Stock Added</option>
                                <option value="low-stock">Low Stock</option>
                                <option value="expiry">Expiry Report</option>
                            </select>
                        </div>
                        
                        <div id="monthly-options" style="display: none;">
                            <div class="form-group">
                                <label for="sales-month">Month</label>
                                <select id="sales-month">
                                    <option value="January">January</option>
                                    <option value="February">February</option>
                                    <option value="March">March</option>
                                    <option value="April">April</option>
                                    <option value="May">May</option>
                                    <option value="June">June</option>
                                    <option value="July">July</option>
                                    <option value="August">August</option>
                                    <option value="September">September</option>
                                    <option value="October">October</option>
                                    <option value="November">November</option>
                                    <option value="December">December</option>
                                </select>
                            </div>
                            
                            <div class="form-group">
                                <label for="sales-year">Year</label>
                                <input type="number" id="sales-year" min="2020" max="2030" value="2023">
                            </div>
                        </div>
                        
                        <button id="generate-sales-report-btn">Generate Report</button>
                        
                        <div id="sales-report-container" style="margin-top: 20px;">
                            <!-- Report will be displayed here -->
                        </div>
                    </div>
                </div>
            </div>
        </div>

        <!-- Modals -->
        <div id="employee-modal" class="modal">
            <div class="modal-content">
                <span class="close-modal">&times;</span>
                <h3>Add Employee</h3>
                <form id="employee-form">
                    <div class="form-group">
                        <label for="full-name">Full Name</label>
                        <input type="text" id="full-name" required>
                    </div>
                    
                    <div class="form-group">
                        <label for="address">Address</label>
                        <textarea id="address"></textarea>
                    </div>
                    
                    <div class="form-group">
                        <label for="contact">Contact</label>
                        <input type="text" id="contact">
                    </div>
                    
                    <div class="form-group">
                        <label for="join-date">Join Date</label>
                        <input type="date" id="join-date" required>
                    </div>
                    
                    <div class="form-group">
                        <label for="salary">Salary (Rs.)</label>
                        <input type="number" id="salary" min="0" step="0.01" required>
                    </div>
                    
                    <div class="form-group">
                        <label for="overtime-rate">Overtime Rate (Rs./hour)</label>
                        <input type="number" id="overtime-rate" min="0" step="0.01" value="0">
                    </div>
                    
                    <div class="form-group">
                        <label for="deductions">Deductions (Rs.)</label>
                        <input type="number" id="deductions" min="0" step="0.01" value="0">
                    </div>
                    
                    <button type="submit">Save Employee</button>
                </form>
            </div>
        </div>

        <div id="income-modal" class="modal">
            <div class="modal-content">
                <span class="close-modal">&times;</span>
                <h3>Add Income</h3>
                <form id="income-form">
                    <div class="form-group">
                        <label for="income-source">Source</label>
                        <input type="text" id="income-source" required>
                    </div>
                    
                    <div class="form-group">
                        <label for="income-amount">Amount (Rs.)</label>
                        <input type="number" id="income-amount" min="0" step="0.01" required>
                    </div>
                    
                    <div class="form-group">
                        <label for="income-date">Date</label>
                        <input type="date" id="income-date" required>
                    </div>
                    
                    <div class="form-group">
                        <label for="income-description">Description</label>
                        <textarea id="income-description"></textarea>
                    </div>
                    
                    <button type="submit">Save Income</button>
                </form>
            </div>
        </div>

        <div id="expense-modal" class="modal">
            <div class="modal-content">
                <span class="close-modal">&times;</span>
                <h3>Add Expense</h3>
                <form id="expense-form">
                    <div class="form-group">
                        <label for="expense-category">Category</label>
                        <select id="expense-category" required>
                            <option value="">Select Category</option>
                            <option value="Rent">Rent</option>
                            <option value="Utilities">Utilities</option>
                            <option value="Salaries">Salaries</option>
                            <option value="Supplies">Supplies</option>
                            <option value="Marketing">Marketing</option>
                            <option value="Other">Other</option>
                        </select>
                    </div>
                    
                    <div class="form-group">
                        <label for="expense-amount">Amount (Rs.)</label>
                        <input type="number" id="expense-amount" min="0" step="0.01" required>
                    </div>
                    
                    <div class="form-group">
                        <label for="expense-date">Date</label>
                        <input type="date" id="expense-date" required>
                    </div>
                    
                    <div class="form-group">
                        <label for="expense-description">Description</label>
                        <textarea id="expense-description"></textarea>
                    </div>
                    
                    <button type="submit">Save Expense</button>
                </form>
            </div>
        </div>

        <div id="product-modal" class="modal">
            <div class="modal-content">
                <span class="close-modal">&times;</span>
                <h3>Add Product</h3>
                <form id="product-form">
                    <div class="form-group">
                        <label for="product-name">Product Name</label>
                        <input type="text" id="product-name" required>
                    </div>
                    
                    <div class="form-group">
                        <label for="product-brand">Brand</label>
                        <input type="text" id="product-brand">
                    </div>
                    
                    <div class="form-group">
                        <label for="product-category">Category</label>
                        <input type="text" id="product-category">
                    </div>
                    
                    <div class="form-group">
                        <label for="weight-volume">Weight/Volume</label>
                        <input type="text" id="weight-volume">
                    </div>
                    
                    <div class="form-group">
                        <label for="packet-count">Packet Count</label>
                        <input type="number" id="packet-count" min="1" value="1">
                    </div>
                    
                    <div class="form-group">
                        <label for="product-price">Price (Rs.)</label>
                        <input type="number" id="product-price" min="0" step="0.01" required>
                    </div>
                    
                    <div class="form-group">
                        <label for="alternate-prices">Alternate Prices (JSON)</label>
                        <textarea id="alternate-prices"></textarea>
                    </div>
                    
                    <div class="form-group">
                        <label for="product-comments">Comments</label>
                        <textarea id="product-comments"></textarea>
                    </div>
                    
                    <div class="form-group">
                        <label for="expiry-date">Expiry Date</label>
                        <input type="date" id="expiry-date">
                    </div>
                    
                    <button type="submit">Save Product</button>
                </form>
            </div>
        </div>

        <div id="notification" class="notification"></div>
    </div>

    <script>
        // Theme toggle functionality
        document.getElementById('theme-toggle').addEventListener('click', function() {
            document.body.classList.toggle('dark-mode');
            const icon = this.querySelector('i');
            if (document.body.classList.contains('dark-mode')) {
                icon.classList.remove('fa-moon');
                icon.classList.add('fa-sun');
            } else {
                icon.classList.remove('fa-sun');
                icon.classList.add('fa-moon');
            }
        });

        // DOM elements
        const dashboard = document.getElementById('dashboard');
        const taskModule = document.getElementById('task-module');
        const payrollModule = document.getElementById('payroll-module');
        const financialModule = document.getElementById('financial-module');
        const posModule = document.getElementById('pos-module');
        const aiChatModule = document.getElementById('ai-chat-module');
        
        const taskModuleBtn = document.getElementById('task-module-btn');
        const payrollModuleBtn = document.getElementById('payroll-module-btn');
        const financialModuleBtn = document.getElementById('financial-module-btn');
        const posModuleBtn = document.getElementById('pos-module-btn');
        const aiChatModuleBtn = document.getElementById('ai-chat-module-btn');
        
        const taskBackBtn = document.getElementById('task-back-btn');
        const payrollBackBtn = document.getElementById('payroll-back-btn');
        const financialBackBtn = document.getElementById('financial-back-btn');
        const posBackBtn = document.getElementById('pos-back-btn');
        const aiChatBackBtn = document.getElementById('ai-chat-back-btn');
        
        // Module navigation
        taskModuleBtn.addEventListener('click', () => {
            dashboard.style.display = 'none';
            taskModule.style.display = 'block';
            loadTasks();
        });
        
        payrollModuleBtn.addEventListener('click', () => {
            dashboard.style.display = 'none';
            payrollModule.style.display = 'block';
            loadEmployees();
            loadEmployeesForPayslip();
        });
        
        financialModuleBtn.addEventListener('click', () => {
            dashboard.style.display = 'none';
            financialModule.style.display = 'block';
            loadIncomeRecords();
            loadExpenseRecords();
            loadFinancialSummary();
        });
        
        posModuleBtn.addEventListener('click', () => {
            dashboard.style.display = 'none';
            posModule.style.display = 'block';
            loadProducts();
        });
        
        aiChatModuleBtn.addEventListener('click', () => {
            dashboard.style.display = 'none';
            aiChatModule.style.display = 'block';
        });
        
        taskBackBtn.addEventListener('click', () => {
            taskModule.style.display = 'none';
            dashboard.style.display = 'block';
        });
        
        payrollBackBtn.addEventListener('click', () => {
            payrollModule.style.display = 'none';
            dashboard.style.display = 'block';
        });
        
        financialBackBtn.addEventListener('click', () => {
            financialModule.style.display = 'none';
            dashboard.style.display = 'block';
        });
        
        posBackBtn.addEventListener('click', () => {
            posModule.style.display = 'none';
            dashboard.style.display = 'block';
        });
        
        aiChatBackBtn.addEventListener('click', () => {
            aiChatModule.style.display = 'none';
            dashboard.style.display = 'block';
        });
        
        // Tab navigation
        document.querySelectorAll('.tab').forEach(tab => {
            tab.addEventListener('click', () => {
                const tabName = tab.getAttribute('data-tab');
                const tabContents = tab.parentElement.parentElement.querySelectorAll('.tab-content');
                
                // Remove active class from all tabs and contents
                tab.parentElement.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
                tabContents.forEach(content => content.classList.remove('active'));
                
                // Add active class to clicked tab and corresponding content
                tab.classList.add('active');
                document.getElementById(`${tabName}-tab`).classList.add('active');
            });
        });
        
        // Modal handling
        const modals = document.querySelectorAll('.modal');
        const closeModalBtns = document.querySelectorAll('.close-modal');
        
        closeModalBtns.forEach(btn => {
            btn.addEventListener('click', () => {
                btn.closest('.modal').style.display = 'none';
            });
        });
        
        window.addEventListener('click', (e) => {
            if (e.target.classList.contains('modal')) {
                e.target.style.display = 'none';
            }
        });
        
        // Notification function
        function showNotification(message, isError = false) {
            const notification = document.getElementById('notification');
            notification.textContent = message;
            notification.className = 'notification show';
            if (isError) {
                notification.classList.add('error');
            }
            
            setTimeout(() => {
                notification.classList.remove('show');
            }, 3000);
        }
        
        // Local storage keys
        const STORAGE_KEYS = {
            TASKS: 'business_tasks',
            EMPLOYEES: 'business_employees',
            INCOME_RECORDS: 'business_income_records',
            EXPENSE_RECORDS: 'business_expense_records',
            PRODUCTS: 'business_products',
            SALES: 'business_sales',
            SALARY_RECORDS: 'business_salary_records'
        };
        
        // Helper functions for local storage
        function getStorageData(key) {
            const data = localStorage.getItem(key);
            return data ? JSON.parse(data) : [];
        }
        
        function setStorageData(key, data) {
            localStorage.setItem(key, JSON.stringify(data));
        }
        
        // Generate unique ID
        function generateId() {
            return Date.now().toString(36) + Math.random().toString(36).substr(2);
        }
        
        // Task Manager Functions
        function loadTasks() {
            const tasks = getStorageData(STORAGE_KEYS.TASKS);
            const pendingTasks = tasks.filter(task => task.status === 'pending');
            const completedTasks = tasks.filter(task => task.status === 'completed');
            
            document.getElementById('pending-count').textContent = pendingTasks.length;
            document.getElementById('completed-count').textContent = completedTasks.length;
            
            renderTasks('pending-tasks', pendingTasks);
            renderTasks('completed-tasks', completedTasks);
        }
        
        function renderTasks(containerId, tasks) {
            const container = document.getElementById(containerId);
            container.innerHTML = '';
            
            if (tasks.length === 0) {
                container.innerHTML = '<p>No tasks found</p>';
                return;
            }
            
            tasks.forEach(task => {
                const taskElement = document.createElement('div');
                taskElement.className = 'task-item';
                
                // Add status class based on due date
                const dueDate = new Date(task.dueDate);
                const today = new Date();
                const diffTime = dueDate - today;
                const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
                
                if (diffDays < 0) {
                    taskElement.classList.add('task-overdue');
                } else if (diffDays <= 3) {
                    taskElement.classList.add('task-due-soon');
                } else {
                    taskElement.classList.add('task-normal');
                }
                
                taskElement.innerHTML = `
                    <div class="task-status status-${task.status}">${task.status}</div>
                    <h4>${task.title}</h4>
                    <div class="task-meta">
                        <div>
                            <label>Category</label>
                            <span>${task.category}</span>
                        </div>
                        <div>
                            <label>Type</label>
                            <span>${task.type}</span>
                        </div>
                        <div>
                            <label>Due Date</label>
                            <span>${new Date(task.dueDate).toLocaleDateString()}</span>
                        </div>
                        <div>
                            <label>Cost</label>
                            <span>Rs. ${task.cost}</span>
                        </div>
                    </div>
                    <div class="task-actions">
                        ${task.status === 'pending' ? 
                            `<button class="btn-success" onclick="updateTaskStatus('${task.id}', 'completed')">Complete</button>` : 
                            `<button class="btn-secondary" onclick="updateTaskStatus('${task.id}', 'pending')">Reopen</button>`
                        }
                        <button class="btn-danger" onclick="deleteTask('${task.id}')">Delete</button>
                    </div>
                `;
                
                container.appendChild(taskElement);
            });
        }
        
        // Task form submission
        document.getElementById('task-form').addEventListener('submit', function(e) {
            e.preventDefault();
            
            const taskData = {
                id: generateId(),
                title: document.getElementById('task-name').value,
                category: document.getElementById('task-category').value,
                type: document.getElementById('task-type').value,
                cost: parseFloat(document.getElementById('task-cost').value) || 0,
                dueDate: document.getElementById('task-due-date').value,
                status: 'pending',
                comments: [{
                    text: document.getElementById('task-comment').value,
                    createdAt: new Date().toISOString()
                }],
                createdAt: new Date().toISOString()
            };
            
            const tasks = getStorageData(STORAGE_KEYS.TASKS);
            tasks.push(taskData);
            setStorageData(STORAGE_KEYS.TASKS, tasks);
            
            showNotification('Task added successfully!');
            this.reset();
            loadTasks();
        });
        
        // Update task status
        function updateTaskStatus(taskId, status) {
            const tasks = getStorageData(STORAGE_KEYS.TASKS);
            const taskIndex = tasks.findIndex(task => task.id === taskId);
            
            if (taskIndex !== -1) {
                tasks[taskIndex].status = status;
                setStorageData(STORAGE_KEYS.TASKS, tasks);
                loadTasks();
                showNotification(`Task marked as ${status}!`);
            }
        }
        
        // Delete task
        function deleteTask(taskId) {
            if (!confirm('Are you sure you want to delete this task?')) return;
            
            const tasks = getStorageData(STORAGE_KEYS.TASKS);
            const filteredTasks = tasks.filter(task => task.id !== taskId);
            setStorageData(STORAGE_KEYS.TASKS, filteredTasks);
            loadTasks();
            showNotification('Task deleted successfully!');
        }
        
        // Payroll Functions
        function loadEmployees() {
            const employees = getStorageData(STORAGE_KEYS.EMPLOYEES);
            const tbody = document.getElementById('employees-table-body');
            tbody.innerHTML = '';
            
            employees.forEach(employee => {
                const row = document.createElement('tr');
                row.innerHTML = `
                    <td>${employee.id}</td>
                    <td>${employee.name}</td>
                    <td>${employee.contact || 'N/A'}</td>
                    <td>Rs. ${employee.salary}</td>
                    <td>${new Date(employee.joinDate).toLocaleDateString()}</td>
                    <td>
                        <button class="btn-secondary" onclick="editEmployee('${employee.id}')">Edit</button>
                        <button class="btn-danger" onclick="deleteEmployee('${employee.id}')">Delete</button>
                    </td>
                `;
                tbody.appendChild(row);
            });
        }
        
        function loadEmployeesForPayslip() {
            const employees = getStorageData(STORAGE_KEYS.EMPLOYEES);
            const select = document.getElementById('employee-select');
            select.innerHTML = '<option value="">Select Employee</option>';
            
            employees.forEach(employee => {
                const option = document.createElement('option');
                option.value = employee.id;
                option.textContent = employee.name;
                select.appendChild(option);
            });
        }
        
        // Add employee button
        document.getElementById('add-employee-btn').addEventListener('click', () => {
            document.getElementById('employee-modal').style.display = 'block';
        });
        
        // Employee form submission
        document.getElementById('employee-form').addEventListener('submit', function(e) {
            e.preventDefault();
            
            const employeeData = {
                id: generateId(),
                name: document.getElementById('full-name').value,
                address: document.getElementById('address').value,
                contact: document.getElementById('contact').value,
                joinDate: document.getElementById('join-date').value,
                salary: parseFloat(document.getElementById('salary').value) || 0,
                overtimeRate: parseFloat(document.getElementById('overtime-rate').value) || 0,
                deductions: parseFloat(document.getElementById('deductions').value) || 0,
                createdAt: new Date().toISOString()
            };
            
            const employees = getStorageData(STORAGE_KEYS.EMPLOYEES);
            employees.push(employeeData);
            setStorageData(STORAGE_KEYS.EMPLOYEES, employees);
            
            showNotification('Employee added successfully!');
            this.reset();
            document.getElementById('employee-modal').style.display = 'none';
            loadEmployees();
            loadEmployeesForPayslip();
        });
        
        // Payslip form submission
        document.getElementById('payslip-form').addEventListener('submit', function(e) {
            e.preventDefault();
            
            const employeeId = document.getElementById('employee-select').value;
            const month = document.getElementById('month').value;
            const year = parseInt(document.getElementById('year').value);
            const overtimeHours = parseFloat(document.getElementById('overtime-hours').value) || 0;
            const deductions = parseFloat(document.getElementById('deductions').value) || 0;
            
            const employees = getStorageData(STORAGE_KEYS.EMPLOYEES);
            const employee = employees.find(emp => emp.id === employeeId);
            
            if (!employee) {
                showNotification('Employee not found!', true);
                return;
            }
            
            // Calculate payslip
            const basicSalary = employee.salary;
            const overtimePay = overtimeHours * employee.overtimeRate;
            const grossSalary = basicSalary + overtimePay;
            const netSalary = grossSalary - deductions;
            
            const payslipData = {
                id: generateId(),
                employeeId,
                employeeName: employee.name,
                month,
                year,
                basicSalary,
                overtimeHours,
                overtimePay,
                deductions,
                grossSalary,
                netSalary,
                createdAt: new Date().toISOString()
            };
            
            const salaryRecords = getStorageData(STORAGE_KEYS.SALARY_RECORDS);
            salaryRecords.push(payslipData);
            setStorageData(STORAGE_KEYS.SALARY_RECORDS, salaryRecords);
            
            showNotification('Payslip generated successfully!');
            this.reset();
            
            // Display payslip
            displayPayslip(payslipData);
        });
        
        function displayPayslip(payslip) {
            const container = document.getElementById('payslip-form').parentElement;
            
            const payslipHtml = `
                <div class="invoice-preview">
                    <div class="invoice-header">
                        <div class="invoice-title">Payslip</div>
                        <div>Month: ${payslip.month} ${payslip.year}</div>
                    </div>
                    
                    <div class="invoice-meta">
                        <div>
                            <strong>Employee:</strong> ${payslip.employeeName}
                        </div>
                    </div>
                    
                    <div class="invoice-details">
                        <div class="invoice-row">
                            <span>Basic Salary:</span>
                            <span>Rs. ${payslip.basicSalary.toFixed(2)}</span>
                        </div>
                        <div class="invoice-row">
                            <span>Overtime Pay:</span>
                            <span>Rs. ${payslip.overtimePay.toFixed(2)}</span>
                        </div>
                        <div class="invoice-row">
                            <span>Deductions:</span>
                            <span>Rs. ${payslip.deductions.toFixed(2)}</span>
                        </div>
                        <div class="invoice-row total">
                            <span>Net Salary:</span>
                            <span>Rs. ${payslip.netSalary.toFixed(2)}</span>
                        </div>
                    </div>
                </div>
            `;
            
            container.insertAdjacentHTML('beforeend', payslipHtml);
        }
        
        // Financial Functions
        function loadIncomeRecords() {
            const incomeRecords = getStorageData(STORAGE_KEYS.INCOME_RECORDS);
            const tbody = document.getElementById('income-table-body');
            tbody.innerHTML = '';
            
            incomeRecords.forEach(record => {
                const row = document.createElement('tr');
                row.innerHTML = `
                    <td>${record.source}</td>
                    <td>Rs. ${record.amount.toFixed(2)}</td>
                    <td>${new Date(record.date).toLocaleDateString()}</td>
                    <td>${record.description || 'N/A'}</td>
                    <td>
                        <button class="btn-secondary" onclick="editIncome('${record.id}')">Edit</button>
                        <button class="btn-danger" onclick="deleteIncome('${record.id}')">Delete</button>
                    </td>
                `;
                tbody.appendChild(row);
            });
        }
        
        function loadExpenseRecords() {
            const expenseRecords = getStorageData(STORAGE_KEYS.EXPENSE_RECORDS);
            const tbody = document.getElementById('expense-table-body');
            tbody.innerHTML = '';
            
            expenseRecords.forEach(record => {
                const row = document.createElement('tr');
                row.innerHTML = `
                    <td>${record.category}</td>
                    <td>Rs. ${record.amount.toFixed(2)}</td>
                    <td>${new Date(record.date).toLocaleDateString()}</td>
                    <td>${record.description || 'N/A'}</td>
                    <td>
                        <button class="btn-secondary" onclick="editExpense('${record.id}')">Edit</button>
                        <button class="btn-danger" onclick="deleteExpense('${record.id}')">Delete</button>
                    </td>
                `;
                tbody.appendChild(row);
            });
        }
        
        function loadFinancialSummary() {
            const period = document.getElementById('summary-period').value;
            const incomeRecords = getStorageData(STORAGE_KEYS.INCOME_RECORDS);
            const expenseRecords = getStorageData(STORAGE_KEYS.EXPENSE_RECORDS);
            
            // Filter records based on period
            const now = new Date();
            let startDate, endDate;
            
            switch (period) {
                case 'today':
                    startDate = new Date(now.setHours(0, 0, 0, 0));
                    endDate = new Date(now.setHours(23, 59, 59, 999));
                    break;
                case 'month':
                    startDate = new Date(now.getFullYear(), now.getMonth(), 1);
                    endDate = new Date(now.getFullYear(), now.getMonth() + 1, 0);
                    break;
                case 'year':
                    startDate = new Date(now.getFullYear(), 0, 1);
                    endDate = new Date(now.getFullYear(), 11, 31);
                    break;
                default:
                    startDate = new Date(now.setHours(0, 0, 0, 0));
                    endDate = new Date(now.setHours(23, 59, 59, 999));
            }
            
            const filteredIncome = incomeRecords.filter(record => {
                const recordDate = new Date(record.date);
                return recordDate >= startDate && recordDate <= endDate;
            });
            
            const filteredExpenses = expenseRecords.filter(record => {
                const recordDate = new Date(record.date);
                return recordDate >= startDate && recordDate <= endDate;
            });
            
            const totalIncome = filteredIncome.reduce((sum, record) => sum + record.amount, 0);
            const totalExpenses = filteredExpenses.reduce((sum, record) => sum + record.amount, 0);
            const profitLoss = totalIncome - totalExpenses;
            
            document.getElementById('total-income').textContent = `Rs. ${totalIncome.toFixed(2)}`;
            document.getElementById('total-expenses').textContent = `Rs. ${totalExpenses.toFixed(2)}`;
            document.getElementById('profit-loss').textContent = `Rs. ${profitLoss.toFixed(2)}`;
            
            // Create chart
            const ctx = document.getElementById('financial-chart').getContext('2d');
            
            new Chart(ctx, {
                type: 'bar',
                data: {
                    labels: ['Income', 'Expenses'],
                    datasets: [{
                        label: 'Amount (Rs.)',
                        data: [totalIncome, totalExpenses],
                        backgroundColor: [
                            'rgba(16, 185, 129, 0.6)',
                            'rgba(239, 68, 68, 0.6)'
                        ],
                        borderColor: [
                            'rgba(16, 185, 129, 1)',
                            'rgba(239, 68, 68, 1)'
                        ],
                        borderWidth: 1
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    scales: {
                        y: {
                            beginAtZero: true
                        }
                    }
                }
            });
        }
        
        // Add income button
        document.getElementById('add-income-btn').addEventListener('click', () => {
            document.getElementById('income-modal').style.display = 'block';
        });
        
        // Income form submission
        document.getElementById('income-form').addEventListener('submit', function(e) {
            e.preventDefault();
            
            const incomeData = {
                id: generateId(),
                source: document.getElementById('income-source').value,
                amount: parseFloat(document.getElementById('income-amount').value),
                date: document.getElementById('income-date').value,
                description: document.getElementById('income-description').value,
                createdAt: new Date().toISOString()
            };
            
            const incomeRecords = getStorageData(STORAGE_KEYS.INCOME_RECORDS);
            incomeRecords.push(incomeData);
            setStorageData(STORAGE_KEYS.INCOME_RECORDS, incomeRecords);
            
            showNotification('Income record added successfully!');
            this.reset();
            document.getElementById('income-modal').style.display = 'none';
            loadIncomeRecords();
            loadFinancialSummary();
        });
        
        // Add expense button
        document.getElementById('add-expense-btn').addEventListener('click', () => {
            document.getElementById('expense-modal').style.display = 'block';
        });
        
        // Expense form submission
        document.getElementById('expense-form').addEventListener('submit', function(e) {
            e.preventDefault();
            
            const expenseData = {
                id: generateId(),
                category: document.getElementById('expense-category').value,
                amount: parseFloat(document.getElementById('expense-amount').value),
                date: document.getElementById('expense-date').value,
                description: document.getElementById('expense-description').value,
                createdAt: new Date().toISOString()
            };
            
            const expenseRecords = getStorageData(STORAGE_KEYS.EXPENSE_RECORDS);
            expenseRecords.push(expenseData);
            setStorageData(STORAGE_KEYS.EXPENSE_RECORDS, expenseRecords);
            
            showNotification('Expense record added successfully!');
            this.reset();
            document.getElementById('expense-modal').style.display = 'none';
            loadExpenseRecords();
            loadFinancialSummary();
        });
        
        // Refresh summary button
        document.getElementById('refresh-summary-btn').addEventListener('click', loadFinancialSummary);
        
        // POS Functions
        function loadProducts() {
            const products = getStorageData(STORAGE_KEYS.PRODUCTS);
            const tbody = document.getElementById('products-table-body');
            tbody.innerHTML = '';
            
            products.forEach(product => {
                const row = document.createElement('tr');
                row.innerHTML = `
                    <td>${product.id}</td>
                    <td>${product.name}</td>
                    <td>${product.brand || 'N/A'}</td>
                    <td>${product.category}</td>
                    <td>Rs. ${product.price.toFixed(2)}</td>
                    <td>${product.stock}</td>
                    <td>${product.expiryDate ? new Date(product.expiryDate).toLocaleDateString() : 'N/A'}</td>
                    <td>
                        <button class="btn-secondary" onclick="editProduct('${product.id}')">Edit</button>
                        <button class="btn-danger" onclick="deleteProduct('${product.id}')">Delete</button>
                    </td>
                `;
                tbody.appendChild(row);
            });
        }
        
        // Add product button
        document.getElementById('add-product-btn').addEventListener('click', () => {
            document.getElementById('product-modal').style.display = 'block';
        });
        
        // Product form submission
        document.getElementById('product-form').addEventListener('submit', function(e) {
            e.preventDefault();
            
            const productData = {
                id: generateId(),
                name: document.getElementById('product-name').value,
                brand: document.getElementById('product-brand').value,
                category: document.getElementById('product-category').value,
                weightVolume: document.getElementById('weight-volume').value,
                packetCount: parseInt(document.getElementById('packet-count').value) || 1,
                price: parseFloat(document.getElementById('product-price').value),
                comments: document.getElementById('product-comments').value,
                expiryDate: document.getElementById('expiry-date').value,
                stock: 0,
                createdAt: new Date().toISOString()
            };
            
            const products = getStorageData(STORAGE_KEYS.PRODUCTS);
            products.push(productData);
            setStorageData(STORAGE_KEYS.PRODUCTS, products);
            
            showNotification('Product added successfully!');
            this.reset();
            document.getElementById('product-modal').style.display = 'none';
            loadProducts();
        });
        
        // Product search functionality
        let selectedProducts = [];
        
        document.getElementById('product-search').addEventListener('input', function(e) {
            const searchTerm = e.target.value.toLowerCase();
            const searchResults = document.getElementById('product-search-results');
            
            if (searchTerm.length < 2) {
                searchResults.style.display = 'none';
                return;
            }
            
            const products = getStorageData(STORAGE_KEYS.PRODUCTS);
            const filteredProducts = products.filter(product => 
                product.name.toLowerCase().includes(searchTerm) ||
                (product.brand && product.brand.toLowerCase().includes(searchTerm))
            );
            
            searchResults.innerHTML = '';
            
            if (filteredProducts.length === 0) {
                searchResults.innerHTML = '<div class="search-result-item">No products found</div>';
            } else {
                filteredProducts.forEach(product => {
                    const item = document.createElement('div');
                    item.className = 'search-result-item';
                    item.innerHTML = `
                        <strong>${product.name}</strong>
                        <div>Rs. ${product.price.toFixed(2)} | Stock: ${product.stock}</div>
                    `;
                    item.addEventListener('click', () => addProductToSale(product));
                    searchResults.appendChild(item);
                });
            }
            
            searchResults.style.display = 'block';
        });
        
        function addProductToSale(product) {
            // Check if product already in selected products
            const existingProduct = selectedProducts.find(p => p.id === product.id);
            
            if (existingProduct) {
                existingProduct.quantity += 1;
            } else {
                selectedProducts.push({
                    id: product.id,
                    name: product.name,
                    price: product.price,
                    quantity: 1
                });
            }
            
            renderSelectedProducts();
            document.getElementById('product-search').value = '';
            document.getElementById('product-search-results').style.display = 'none';
        }
        
        function renderSelectedProducts() {
            const container = document.getElementById('selected-products-list');
            container.innerHTML = '';
            
            if (selectedProducts.length === 0) {
                container.innerHTML = '<p>No products selected</p>';
                updateSaleSummary();
                return;
            }
            
            selectedProducts.forEach((product, index) => {
                const productRow = document.createElement('div');
                productRow.className = 'product-row';
                productRow.innerHTML = `
                    <div>${product.name}</div>
                    <div>Rs. ${product.price.toFixed(2)}</div>
                    <div>
                        <input type="number" value="${product.quantity}" min="1" onchange="updateProductQuantity(${index}, this.value)">
                    </div>
                    <div>Rs. ${(product.price * product.quantity).toFixed(2)}</div>
                    <div>
                        <button class="btn-danger" onclick="removeProductFromSale(${index})">Remove</button>
                    </div>
                `;
                container.appendChild(productRow);
            });
            
            updateSaleSummary();
        }
        
        function updateProductQuantity(index, quantity) {
            selectedProducts[index].quantity = parseInt(quantity);
            renderSelectedProducts();
        }
        
        function removeProductFromSale(index) {
            selectedProducts.splice(index, 1);
            renderSelectedProducts();
        }
        
        function updateSaleSummary() {
            const subtotal = selectedProducts.reduce((sum, product) => sum + (product.price * product.quantity), 0);
            const discount = parseFloat(document.getElementById('discount').value) || 0;
            const total = subtotal - discount;
            
            document.getElementById('subtotal').textContent = `Rs. ${subtotal.toFixed(2)}`;
            document.getElementById('discount-amount').textContent = `Rs. ${discount.toFixed(2)}`;
            document.getElementById('total-amount').textContent = `Rs. ${total.toFixed(2)}`;
        }
        
        document.getElementById('discount').addEventListener('input', updateSaleSummary);
        
        // Sale form submission
        document.getElementById('sale-form').addEventListener('submit', function(e) {
            e.preventDefault();
            
            if (selectedProducts.length === 0) {
                showNotification('Please add at least one product', true);
                return;
            }
            
            const saleData = {
                id: generateId(),
                items: selectedProducts.map(product => ({
                    productId: product.id,
                    name: product.name,
                    price: product.price,
                    quantity: product.quantity,
                    total: product.price * product.quantity
                })),
                subtotal: selectedProducts.reduce((sum, product) => sum + (product.price * product.quantity), 0),
                discount: parseFloat(document.getElementById('discount').value) || 0,
                total: selectedProducts.reduce((sum, product) => sum + (product.price * product.quantity), 0) - (parseFloat(document.getElementById('discount').value) || 0),
                paymentMethod: document.getElementById('payment-method').value,
                createdAt: new Date().toISOString()
            };
            
            const sales = getStorageData(STORAGE_KEYS.SALES);
            sales.push(saleData);
            setStorageData(STORAGE_KEYS.SALES, sales);
            
            // Update product stock
            const products = getStorageData(STORAGE_KEYS.PRODUCTS);
            selectedProducts.forEach(selectedProduct => {
                const productIndex = products.findIndex(p => p.id === selectedProduct.id);
                if (productIndex !== -1) {
                    products[productIndex].stock -= selectedProduct.quantity;
                }
            });
            setStorageData(STORAGE_KEYS.PRODUCTS, products);
            
            showNotification('Sale completed successfully!');
            this.reset();
            selectedProducts = [];
            renderSelectedProducts();
            loadProducts();
        });
        
        // AI Chat Functions
        const chatMessages = document.getElementById('chat-messages');
        const chatInput = document.getElementById('chat-input');
        const sendChatBtn = document.getElementById('send-chat-btn');
        
        sendChatBtn.addEventListener('click', sendMessage);
        chatInput.addEventListener('keypress', (e) => {
            if (e.key === 'Enter') {
                sendMessage();
            }
        });
        
        function sendMessage() {
            const message = chatInput.value.trim();
            if (!message) return;
            
            addMessageToChat(message, 'user');
            chatInput.value = '';
            
            showTypingIndicator();
            
            setTimeout(() => {
                processUserQuery(message);
            }, 1000);
        }
        
        function addMessageToChat(content, sender) {
            const messageDiv = document.createElement('div');
            messageDiv.className = `message ${sender}-message`;
            
            const messageContent = document.createElement('div');
            messageContent.className = 'message-content';
            messageContent.textContent = content;
            
            messageDiv.appendChild(messageContent);
            chatMessages.appendChild(messageDiv);
            
            chatMessages.scrollTop = chatMessages.scrollHeight;
        }
        
        function showTypingIndicator() {
            const typingDiv = document.createElement('div');
            typingDiv.className = 'message ai-message';
            typingDiv.id = 'typing-indicator';
            
            const typingContent = document.createElement('div');
            typingContent.className = 'typing-indicator';
            typingContent.innerHTML = `
                <span></span>
                <span></span>
                <span></span>
            `;
            
            typingDiv.appendChild(typingContent);
            chatMessages.appendChild(typingDiv);
            
            chatMessages.scrollTop = chatMessages.scrollHeight;
        }
        
        function removeTypingIndicator() {
            const typingIndicator = document.getElementById('typing-indicator');
            if (typingIndicator) {
                typingIndicator.remove();
            }
        }
        
        function processUserQuery(query) {
            removeTypingIndicator();
            
            const lowerQuery = query.toLowerCase();
            let response = '';
            
            // Process different types of queries
            if (lowerQuery.includes('today') && lowerQuery.includes('sales')) {
                response = getTodaySalesSummary();
            } else if (lowerQuery.includes('month') && lowerQuery.includes('sales') && lowerQuery.includes('chart')) {
                response = 'I can generate a monthly sales chart for you. Please go to the POS module and select the Monthly Sales report option.';
            } else if (lowerQuery.includes('month') && lowerQuery.includes('payroll')) {
                response = getMonthlyPayrollCost();
            } else if (lowerQuery.includes('pending') && lowerQuery.includes('task') && lowerQuery.includes('category')) {
                response = getPendingTasksByCategory();
            } else if (lowerQuery.includes('sales') && lowerQuery.includes('expenses') && lowerQuery.includes('chart')) {
                response = 'I can generate a sales vs expenses chart for you. Please go to the Financial module and select the Summary tab.';
            } else if (lowerQuery.includes('income') && lowerQuery.includes('expense') && lowerQuery.includes('summary')) {
                response = getIncomeExpenseSummary();
            } else if (lowerQuery.includes('product') && lowerQuery.includes('sales')) {
                response = getProductSales();
            } else if (lowerQuery.includes('low') && lowerQuery.includes('stock')) {
                response = getLowStockProducts();
            } else if (lowerQuery.includes('employee') && lowerQuery.includes('count')) {
                response = getEmployeeCount();
            } else if (lowerQuery.includes('task') && lowerQuery.includes('count')) {
                response = getTaskCount();
            } else {
                response = "I'm sorry, I didn't understand your query. You can ask me about sales, expenses, payroll, tasks, or products. For example, you can ask 'Show me today's sales summary' or 'Generate a chart of expenses by category'.";
            }
            
            addMessageToChat(response, 'ai');
        }
        
        function getTodaySalesSummary() {
            const sales = getStorageData(STORAGE_KEYS.SALES);
            const today = new Date().toDateString();
            
            const todaySales = sales.filter(sale => {
                const saleDate = new Date(sale.createdAt).toDateString();
                return saleDate === today;
            });
            
            const totalSales = todaySales.reduce((sum, sale) => sum + sale.total, 0);
            const totalItems = todaySales.reduce((sum, sale) => sum + sale.items.length, 0);
            
            return `Today's Sales Summary:\n\nTotal Sales: Rs. ${totalSales.toFixed(2)}\nTotal Items Sold: ${totalItems}\nNumber of Transactions: ${todaySales.length}`;
        }
        
        function getMonthlyPayrollCost() {
            const salaryRecords = getStorageData(STORAGE_KEYS.SALARY_RECORDS);
            const currentMonth = new Date().toLocaleString('default', { month: 'long' });
            const currentYear = new Date().getFullYear();
            
            const monthlyRecords = salaryRecords.filter(record => 
                record.month === currentMonth && record.year === currentYear
            );
            
            const totalPayroll = monthlyRecords.reduce((sum, record) => sum + record.netSalary, 0);
            
            return `Monthly Payroll Cost for ${currentMonth} ${currentYear}:\n\nTotal Payroll: Rs. ${totalPayroll.toFixed(2)}\nNumber of Employees: ${monthlyRecords.length}`;
        }
        
        function getPendingTasksByCategory() {
            const tasks = getStorageData(STORAGE_KEYS.TASKS);
            const pendingTasks = tasks.filter(task => task.status === 'pending');
            
            // Group tasks by category
            const tasksByCategory = {};
            pendingTasks.forEach(task => {
                if (!tasksByCategory[task.category]) {
                    tasksByCategory[task.category] = 0;
                }
                tasksByCategory[task.category] += 1;
            });
            
            let result = 'Pending Tasks by Category:\n\n';
            Object.keys(tasksByCategory).forEach(category => {
                result += `${category}: ${tasksByCategory[category]}\n`;
            });
            
            return result;
        }
        
        function getIncomeExpenseSummary() {
            const incomeRecords = getStorageData(STORAGE_KEYS.INCOME_RECORDS);
            const expenseRecords = getStorageData(STORAGE_KEYS.EXPENSE_RECORDS);
            const currentMonth = new Date().toLocaleString('default', { month: 'long' });
            const currentYear = new Date().getFullYear();
            
            const monthlyIncome = incomeRecords.filter(record => {
                const recordDate = new Date(record.date);
                const recordMonth = recordDate.toLocaleString('default', { month: 'long' });
                return recordMonth === currentMonth && recordDate.getFullYear() === currentYear;
            });
            
            const monthlyExpenses = expenseRecords.filter(record => {
                const recordDate = new Date(record.date);
                const recordMonth = recordDate.toLocaleString('default', { month: 'long' });
                return recordMonth === currentMonth && recordDate.getFullYear() === currentYear;
            });
            
            const totalIncome = monthlyIncome.reduce((sum, record) => sum + record.amount, 0);
            const totalExpenses = monthlyExpenses.reduce((sum, record) => sum + record.amount, 0);
            const profitLoss = totalIncome - totalExpenses;
            
            return `Income and Expense Summary for ${currentMonth} ${currentYear}:\n\nTotal Income: Rs. ${totalIncome.toFixed(2)}\nTotal Expenses: Rs. ${totalExpenses.toFixed(2)}\nProfit/Loss: Rs. ${profitLoss.toFixed(2)}`;
        }
        
        function getProductSales() {
            const sales = getStorageData(STORAGE_KEYS.SALES);
            const products = getStorageData(STORAGE_KEYS.PRODUCTS);
            
            // Calculate sales by product
            const productSales = {};
            
            sales.forEach(sale => {
                sale.items.forEach(item => {
                    const product = products.find(p => p.id === item.productId);
                    if (product) {
                        if (!productSales[product.name]) {
                            productSales[product.name] = 0;
                        }
                        productSales[product.name] += item.quantity;
                    }
                });
            });
            
            // Sort by quantity sold
            const sortedProducts = Object.keys(productSales)
                .map(name => ({ name, quantity: productSales[name] }))
                .sort((a, b) => b.quantity - a.quantity);
            
            let result = 'Product Sales:\n\n';
            sortedProducts.slice(0, 10).forEach(product => {
                result += `${product.name}: ${product.quantity} units\n`;
            });
            
            return result;
        }
        
        function getLowStockProducts() {
            const products = getStorageData(STORAGE_KEYS.PRODUCTS);
            const lowStockProducts = products.filter(product => product.stock < 10);
            
            let result = 'Low Stock Products (less than 10 units):\n\n';
            lowStockProducts.forEach(product => {
                result += `${product.name}: ${product.stock} units\n`;
            });
            
            return result;
        }
        
        function getEmployeeCount() {
            const employees = getStorageData(STORAGE_KEYS.EMPLOYEES);
            return `Total number of employees: ${employees.length}`;
        }
        
        function getTaskCount() {
            const tasks = getStorageData(STORAGE_KEYS.TASKS);
            const pendingTasks = tasks.filter(task => task.status === 'pending');
            const completedTasks = tasks.filter(task => task.status === 'completed');
            
            return `Task Summary:\n\nTotal Tasks: ${tasks.length}\nPending Tasks: ${pendingTasks.length}\nCompleted Tasks: ${completedTasks.length}`;
        }
        
        // Initialize the application
        document.addEventListener('DOMContentLoaded', () => {
            // Set default date inputs to today
            const today = new Date().toISOString().split('T')[0];
            document.getElementById('task-due-date').value = today;
            document.getElementById('join-date').value = today;
            document.getElementById('income-date').value = today;
            document.getElementById('expense-date').value = today;
            document.getElementById('expiry-date').value = today;
            
            // Initialize storage if empty
            Object.values(STORAGE_KEYS).forEach(key => {
                if (!localStorage.getItem(key)) {
                    localStorage.setItem(key, JSON.stringify([]));
                }
            });
        });
    </script>
</body>
</html>
