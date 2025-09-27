# i_smart.imaz_BMS
Busiuness Management System
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
        
        // AI Chat Module Functions
        
        // Chat functionality
        const chatMessages = document.getElementById('chat-messages');
        const chatInput = document.getElementById('chat-input');
        const sendChatBtn = document.getElementById('send-chat-btn');
        
        // Send message on button click
        sendChatBtn.addEventListener('click', sendMessage);
        
        // Send message on Enter key
        chatInput.addEventListener('keypress', (e) => {
            if (e.key === 'Enter') {
                sendMessage();
            }
        });
        
        function sendMessage() {
            const message = chatInput.value.trim();
            if (!message) return;
            
            // Add user message to chat
            addMessageToChat(message, 'user');
            
            // Clear input
            chatInput.value = '';
            
            // Show typing indicator
            showTypingIndicator();
            
            // Process the message after a short delay
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
            
            // Scroll to bottom
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
            
            // Scroll to bottom
            chatMessages.scrollTop = chatMessages.scrollHeight;
        }
        
        function removeTypingIndicator() {
            const typingIndicator = document.getElementById('typing-indicator');
            if (typingIndicator) {
                typingIndicator.remove();
            }
        }
        
        function processUserQuery(query) {
            // Remove typing indicator
            removeTypingIndicator();
            
            // Convert query to lowercase for easier matching
            const lowerQuery = query.toLowerCase();
            
            // Initialize response object
            let response = {
                type: 'text',
                content: '',
                data: null,
                title: '',
                chartType: null
            };
            
            // Pattern matching for different types of queries
            if (lowerQuery.includes('today') && lowerQuery.includes('sales')) {
                response = getTodaySalesSummary();
            } else if (lowerQuery.includes('month') && lowerQuery.includes('sales') && lowerQuery.includes('chart')) {
                response = getMonthlySalesChart();
            } else if (lowerQuery.includes('month') && lowerQuery.includes('payroll')) {
                response = getMonthlyPayrollCost();
            } else if (lowerQuery.includes('pending') && lowerQuery.includes('task') && lowerQuery.includes('category')) {
                response = getPendingTasksByCategory();
            } else if (lowerQuery.includes('sales') && lowerQuery.includes('expenses') && lowerQuery.includes('chart')) {
                response = getSalesVsExpensesChart();
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
                response = {
                    type: 'text',
                    content: "I'm sorry, I didn't understand your query. You can ask me about sales, expenses, payroll, tasks, or products. For example, you can ask 'Show me today's sales summary' or 'Generate a chart of expenses by category'."
                };
            }
            
            // Add the response to the chat
            addResponseToChat(response);
        }
        
        function addResponseToChat(response) {
            const messageDiv = document.createElement('div');
            messageDiv.className = 'message ai-message';
            
            const messageContent = document.createElement('div');
            messageContent.className = 'message-content';
            
            if (response.type === 'text') {
                messageContent.textContent = response.content;
            } else if (response.type === 'table') {
                messageContent.innerHTML = `
                    <div>${response.content}</div>
                    <div class="chat-response-table-container">
                        <table class="chat-response-table">
                            <thead>
                                <tr>
                                    ${response.columns.map(col => `<th>${col}</th>`).join('')}
                                </tr>
                            </thead>
                            <tbody>
                                ${response.data.map(row => `
                                    <tr>
                                        ${response.columns.map(col => `<td>${row[col]}</td>`).join('')}
                                    </tr>
                                `).join('')}
                            </tbody>
                        </table>
                    </div>
                `;
            } else if (response.type === 'chart') {
                messageContent.innerHTML = `
                    <div>${response.content}</div>
                    <div class="chat-response-chart" id="chart-${Date.now()}"></div>
                `;
                
                // Create chart after adding to DOM
                setTimeout(() => {
                    createChart(`chart-${Date.now()}`, response.chartType, response.data, response.title);
                }, 100);
            }
            
            messageDiv.appendChild(messageContent);
            
            // Add export buttons if applicable
            if (response.type !== 'text') {
                const messageActions = document.createElement('div');
                messageActions.className = 'message-actions';
                
                const exportExcelBtn = document.createElement('button');
                exportExcelBtn.className = 'btn-secondary';
                exportExcelBtn.textContent = 'Export to Excel';
                exportExcelBtn.addEventListener('click', () => exportResponseToExcel(response));
                
                const exportPdfBtn = document.createElement('button');
                exportPdfBtn.className = 'btn-secondary';
                exportPdfBtn.textContent = 'Export to PDF';
                exportPdfBtn.addEventListener('click', () => exportResponseToPdf(response));
                
                messageActions.appendChild(exportExcelBtn);
                messageActions.appendChild(exportPdfBtn);
                messageDiv.appendChild(messageActions);
            }
            
            chatMessages.appendChild(messageDiv);
            
            // Scroll to bottom
            chatMessages.scrollTop = chatMessages.scrollHeight;
        }
        
        function createChart(canvasId, chartType, data, title) {
            const canvas = document.getElementById(canvasId);
            if (!canvas) return;
            
            const ctx = canvas.getContext('2d');
            
            let chart;
            
            if (chartType === 'bar') {
                chart = new Chart(ctx, {
                    type: 'bar',
                    data: data,
                    options: {
                        responsive: true,
                        maintainAspectRatio: false,
                        plugins: {
                            title: {
                                display: true,
                                text: title
                            }
                        }
                    }
                });
            } else if (chartType === 'pie') {
                chart = new Chart(ctx, {
                    type: 'pie',
                    data: data,
                    options: {
                        responsive: true,
                        maintainAspectRatio: false,
                        plugins: {
                            title: {
                                display: true,
                                text: title
                            }
                        }
                    }
                });
            } else if (chartType === 'line') {
                chart = new Chart(ctx, {
                    type: 'line',
                    data: data,
                    options: {
                        responsive: true,
                        maintainAspectRatio: false,
                        plugins: {
                            title: {
                                display: true,
                                text: title
                            }
                        }
                    }
                });
            }
        }
        
        function exportResponseToExcel(response) {
            const wb = XLSX.utils.book_new();
            
            if (response.type === 'table') {
                const ws = XLSX.utils.json_to_sheet(response.data);
                XLSX.utils.book_append_sheet(wb, ws, 'Report');
            } else if (response.type === 'chart') {
                // For charts, we'll create a simple data table
                const ws = XLSX.utils.json_to_sheet(response.data.datasets[0].data.map((value, index) => ({
                    [response.data.labels[index]]: value
                })));
                XLSX.utils.book_append_sheet(wb, ws, 'Chart Data');
            }
            
            XLSX.writeFile(wb, `report_${Date.now().toString().slice(-6)}.xlsx`);
            showNotification('Report exported to Excel successfully');
        }
        
        function exportResponseToPdf(response) {
            const { jsPDF } = window.jspdf;
            const doc = new jsPDF();
            
            // Add title
            doc.setFontSize(18);
            doc.text(response.title || 'Business Report', 105, 20, { align: 'center' });
            
            // Add generation date
            doc.setFontSize(12);
            doc.text(`Generated on: ${new Date().toLocaleString()}`, 105, 30, { align: 'center' });
            
            // Add content
            let yPosition = 50;
            
            if (response.type === 'text') {
                doc.text(response.content, 20, yPosition);
            } else if (response.type === 'table') {
                // Add table
                doc.text('Data Table:', 20, yPosition);
                yPosition += 10;
                
                // Add headers
                response.columns.forEach((col, index) => {
                    doc.text(col, 20 + (index * 80), yPosition);
                });
                yPosition += 10;
                
                // Add rows
                response.data.forEach(row => {
                    response.columns.forEach((col, index) => {
                        doc.text(row[col].toString(), 20 + (index * 80), yPosition);
                    });
                    yPosition += 10;
                });
            } else if (response.type === 'chart') {
                doc.text('Chart Data:', 20, yPosition);
                yPosition += 10;
                
                // Add labels and values
                response.data.labels.forEach((label, index) => {
                    doc.text(`${label}: ${response.data.datasets[0].data[index]}`, 20, yPosition);
                    yPosition += 10;
                });
            }
            
            doc.save(`report_${Date.now().toString().slice(-6)}.pdf`);
            showNotification('Report exported to PDF successfully');
        }
        
        // Query processing functions
        function getTodaySalesSummary() {
            const sales = getStorageData(STORAGE_KEYS.SALES);
            const today = new Date().toDateString();
            
            const todaySales = sales.filter(sale => {
                const saleDate = new Date(sale.saleDate).toDateString();
                return saleDate === today;
            });
            
            const totalSales = todaySales.reduce((sum, sale) => sum + sale.total, 0);
            const totalItems = todaySales.reduce((sum, sale) => sum + sale.items.length, 0);
            
            return {
                type: 'text',
                content: `Today's Sales Summary:\n\nTotal Sales: Rs. ${totalSales.toFixed(2)}\nTotal Items Sold: ${totalItems}\nNumber of Transactions: ${todaySales.length}`
            };
        }
        
        function getMonthlySalesChart() {
            const sales = getStorageData(STORAGE_KEYS.SALES);
            const currentMonth = new Date().toLocaleString('default', { month: 'long' });
            const currentYear = new Date().getFullYear();
            
            const monthlySales = sales.filter(sale => {
                const saleDate = new Date(sale.saleDate);
                const saleMonth = saleDate.toLocaleString('default', { month: 'long' });
                return saleMonth === currentMonth && saleDate.getFullYear() === currentYear;
            });
            
            // Group sales by day
            const salesByDay = {};
            monthlySales.forEach(sale => {
                const day = new Date(sale.saleDate).getDate();
                if (!salesByDay[day]) {
                    salesByDay[day] = 0;
                }
                salesByDay[day] += sale.total;
            });
            
            const labels = Object.keys(salesByDay).sort((a, b) => a - b);
            const data = labels.map(day => salesByDay[day]);
            
            return {
                type: 'chart',
                chartType: 'bar',
                title: `${currentMonth} ${currentYear} Daily Sales`,
                content: `Here's a chart showing daily sales for ${currentMonth} ${currentYear}:`,
                data: {
                    labels: labels.map(day => `${day} ${currentMonth.slice(0, 3)}`),
                    datasets: [{
                        label: 'Sales (Rs.)',
                        data: data,
                        backgroundColor: 'rgba(37, 99, 235, 0.6)',
                        borderColor: 'rgba(37, 99, 235, 1)',
                        borderWidth: 1
                    }]
                }
            };
        }
        
        function getMonthlyPayrollCost() {
            const employees = getStorageData(STORAGE_KEYS.EMPLOYEES);
            const salaryRecords = getStorageData(STORAGE_KEYS.SALARY_RECORDS);
            const currentMonth = new Date().toLocaleString('default', { month: 'long' });
            const currentYear = new Date().getFullYear();
            
            const monthlyRecords = salaryRecords.filter(record => 
                record.month === currentMonth && record.year === currentYear
            );
            
            const totalPayroll = monthlyRecords.reduce((sum, record) => sum + record.netSalary, 0);
            
            return {
                type: 'text',
                content: `Monthly Payroll Cost for ${currentMonth} ${currentYear}:\n\nTotal Payroll: Rs. ${totalPayroll.toFixed(2)}\nNumber of Employees: ${monthlyRecords.length}`
            };
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
            
            return {
                type: 'table',
                content: 'Pending Tasks by Category:',
                columns: ['Category', 'Count'],
                data: Object.keys(tasksByCategory).map(category => ({
                    'Category': category,
                    'Count': tasksByCategory[category]
                }))
            };
        }
        
        function getSalesVsExpensesChart() {
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
            
            return {
                type: 'chart',
                chartType: 'bar',
                title: `Sales vs Expenses for ${currentMonth} ${currentYear}`,
                content: `Here's a comparison of sales and expenses for ${currentMonth} ${currentYear}:`,
                data: {
                    labels: ['Sales', 'Expenses'],
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
                }
            };
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
            
            return {
                type: 'text',
                content: `Income and Expense Summary for ${currentMonth} ${currentYear}:\n\nTotal Income: Rs. ${totalIncome.toFixed(2)}\nTotal Expenses: Rs. ${totalExpenses.toFixed(2)}\nProfit/Loss: Rs. ${profitLoss.toFixed(2)}`
            };
        }
        
        function getProductSales() {
            const sales = getStorageData(STORAGE_KEYS.SALES);
            const products = getStorageData(STORAGE_KEYS.PRODUCTS);
            
            // Calculate sales by product
            const productSales = {};
            
            sales.forEach(sale => {
                sale.items.forEach(item => {
                    const product = products.find(p => p.id === item.id);
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
            
            return {
                type: 'table',
                content: 'Product Sales:',
                columns: ['Product Name', 'Quantity Sold'],
                data: sortedProducts.slice(0, 10) // Top 10 products
            };
        }
        
        function getLowStockProducts() {
            const products = getStorageData(STORAGE_KEYS.PRODUCTS);
            const lowStockProducts = products.filter(product => (product.stock || 0) < 10);
            
            return {
                type: 'table',
                content: 'Low Stock Products (less than 10 units):',
                columns: ['Product Name', 'Stock'],
                data: lowStockProducts.map(product => ({
                    'Product Name': product.name,
                    'Stock': product.stock || 0
                }))
            };
        }
        
        function getEmployeeCount() {
            const employees = getStorageData(STORAGE_KEYS.EMPLOYEES);
            
            return {
                type: 'text',
                content: `Total number of employees: ${employees.length}`
            };
        }
        
        function getTaskCount() {
            const tasks = getStorageData(STORAGE_KEYS.TASKS);
            const pendingTasks = tasks.filter(task => task.status === 'pending');
            const completedTasks = tasks.filter(task => task.status === 'completed');
            
            return {
                type: 'text',
                content: `Task Summary:\n\nTotal Tasks: ${tasks.length}\nPending Tasks: ${pendingTasks.length}\nCompleted Tasks: ${completedTasks.length}`
            };
        }
        
        // Initialize the application
        document.addEventListener('DOMContentLoaded', () => {
            // Set default date inputs to today
            const today = new Date().toISOString().split('T')[0];
            document.getElementById('task-due-date').value = today;
            document.getElementById('join-date').value = today;
            document.getElementById('income-date').value = today;
            document.getElementById('expense-date').value = today;
            document.getElementById('start-date').value = today;
            document.getElementById('end-date').value = today;
            document.getElementById('expiry-date').value = today;
        });
    </script>
</body>
</html>
