<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Janak Positioning & Surveying - Stock Management</title>
    
    <!-- Tailwind CSS -->
    <script src="https://cdn.tailwindcss.com"></script>
    
    <!-- Google Fonts -->
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap" rel="stylesheet">
    
    <!-- Phosphor Icons -->
    <script src="https://unpkg.com/@phosphor-icons/web"></script>

    <!-- SheetJS for Excel (.xlsx) Export -->
    <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
    
    <style>
        body {
            font-family: 'Inter', sans-serif;
            margin: 0;
            padding: 0;
            background-color: #f8fafc;
        }
        
        /* Custom Scrollbar for the Main Content Area */
        ::-webkit-scrollbar { width: 8px; height: 8px; }
        ::-webkit-scrollbar-track { background: #f1f5f9; border-radius: 4px; }
        ::-webkit-scrollbar-thumb { background: #cbd5e1; border-radius: 4px; }
        ::-webkit-scrollbar-thumb:hover { background: #94a3b8; }

        /* Print Styles */
        @media print {
            body { background-color: white; overflow: auto; height: auto; }
            #github-theme-override, #app-container, #toast-container, .no-print, #edit-modal, #camera-modal, #approval-details-modal, #gatepass-preview-modal, #csv-import-modal, #track-modal, #dispatch-details-modal, #manage-models-modal, #drive-sync-modal, #model-pin-modal, #user-edit-modal, #forgot-password-modal {
                display: none !important;
            }
            #gatepass-print-view {
                display: block !important;
                position: absolute;
                top: 0;
                left: 0;
                width: 100%;
                margin: 0;
                padding: 20px;
                background-color: white;
            }
        }
        
        @media screen {
            #gatepass-print-view {
                display: none;
            }
        }
    </style>
</head>
<body>

    <!-- GITHUB THEME OVERRIDE WRAPPER: Fixed to exactly viewport size so bottom is never cut off -->
    <div id="github-theme-override" class="fixed top-0 left-0 w-full h-[100vh] h-[100dvh] z-[999999] bg-slate-50 overflow-hidden flex flex-col md:flex-row text-slate-800 antialiased">

        <!-- Login Screen -->
        <div id="login-screen" class="absolute inset-0 z-[100] flex items-center justify-center bg-slate-900 px-4">
            <div class="bg-white p-8 rounded-2xl shadow-2xl max-w-sm w-full">
                <div class="flex justify-center mb-5">
                    <img src="Janak_Logo.png" alt="Janak Positioning & Surveying Systems Pvt Ltd" class="h-16 w-auto object-contain">
                </div>
                <h2 class="text-2xl font-bold text-center text-slate-800 mb-2">Janak Stock App</h2>
                <p class="text-center text-slate-500 text-sm mb-6">Sign in to your account</p>
                
                <div id="login-error" class="hidden mb-4 p-3 bg-rose-50 text-rose-600 text-sm rounded-lg border border-rose-200 text-center font-medium">
                    Invalid username or password
                </div>

                <!-- Quick Role Select Buttons -->
                <div class="mb-5">
                    <p class="text-xs font-semibold text-slate-400 uppercase tracking-wider text-center mb-2.5">Select Role / Quick Login</p>
                    <div class="grid grid-cols-2 gap-2.5">
                        <button type="button" id="role-btn-admin" onclick="selectLoginRole('admin')" class="py-2.5 px-3 bg-slate-100 hover:bg-slate-200 text-slate-700 border border-slate-200 text-xs font-bold rounded-xl transition flex items-center justify-center gap-1.5 shadow-sm">
                            <i class="ph-fill ph-shield-check text-slate-500 text-base" id="role-icon-admin"></i> Admin
                        </button>
                        <button type="button" id="role-btn-manager" onclick="selectLoginRole('manager')" class="py-2.5 px-3 bg-slate-100 hover:bg-slate-200 text-slate-700 border border-slate-200 text-xs font-bold rounded-xl transition flex items-center justify-center gap-1.5 shadow-sm">
                            <i class="ph-fill ph-user-circle text-slate-500 text-base" id="role-icon-manager"></i> Store Manager
                        </button>
                    </div>
                </div>

                <form id="form-login" onsubmit="handleLogin(event)" class="space-y-4">
                    <div>
                        <label class="block text-sm font-semibold text-slate-700 mb-1">Username</label>
                        <input type="text" id="login-user" required class="w-full px-4 py-3 border border-slate-300 rounded-xl focus:ring-2 focus:ring-blue-500 outline-none">
                    </div>
                    <div>
                        <label class="block text-sm font-semibold text-slate-700 mb-1">Password</label>
                        <input type="password" id="login-pass" required class="w-full px-4 py-3 border border-slate-300 rounded-xl focus:ring-2 focus:ring-blue-500 outline-none">
                    </div>
                    
                    <div class="flex items-center justify-end text-xs">
                        <button type="button" onclick="openForgotPasswordModal()" class="text-blue-600 hover:underline font-semibold">Forgot Password?</button>
                    </div>

                    <button type="submit" class="w-full bg-blue-600 text-white font-bold py-3 px-4 rounded-xl hover:bg-blue-700 transition shadow-md">
                        Login to System
                    </button>
                </form>
            </div>
        </div>

        <!-- App Container -->
        <div id="app-container" class="hidden w-full h-full flex-col md:flex-row relative">
            
            <!-- Sidebar -->
            <aside class="w-full md:w-64 bg-slate-900 text-white flex flex-col shadow-xl z-20 shrink-0 h-full">
                <div class="p-6 border-b border-slate-800 flex items-center gap-4 md:block shrink-0">
                    <div class="bg-white p-2.5 rounded-xl inline-flex items-center justify-center md:mb-3 shrink-0 shadow-inner">
                        <img src="Janak_Logo.png" alt="Janak Positioning & Surveying" class="h-10 w-auto object-contain">
                    </div>
                    <div>
                        <h1 class="font-bold text-base leading-tight hidden md:block">Janak Positioning & <br><span class="text-blue-400 font-medium text-xs">Surveying Systems</span></h1>
                        <h1 class="font-bold text-lg leading-tight md:hidden">Janak Stock App</h1>
                    </div>
                </div>
                
                <!-- Added md:overflow-y-auto so sidebar can scroll if it gets too long -->
                <nav class="flex-1 py-4 md:py-6 px-3 flex md:flex-col gap-2 overflow-x-auto md:overflow-y-auto">
                    <button type="button" onclick="switchTab('dashboard')" id="nav-dashboard" class="nav-btn whitespace-nowrap md:w-full flex items-center gap-3 px-4 py-3 rounded-xl bg-blue-600 text-white font-medium transition-colors text-left shrink-0">
                        <i class="ph ph-squares-four text-xl"></i>
                        <span>Dashboard</span>
                    </button>
                    <button type="button" onclick="switchTab('material-in')" id="nav-material-in" class="nav-btn whitespace-nowrap md:w-full flex items-center gap-3 px-4 py-3 rounded-xl text-slate-400 hover:bg-slate-800 hover:text-white font-medium transition-colors text-left shrink-0">
                        <i class="ph ph-box-arrow-down text-xl"></i>
                        <span>Material In</span>
                    </button>
                    <button type="button" onclick="switchTab('material-out')" id="nav-material-out" class="nav-btn whitespace-nowrap md:w-full flex items-center gap-3 px-4 py-3 rounded-xl text-slate-400 hover:bg-slate-800 hover:text-white font-medium transition-colors text-left shrink-0">
                        <i class="ph ph-truck text-xl"></i>
                        <span>Material Out</span>
                    </button>
                    <button type="button" onclick="switchTab('dispatch-history')" id="nav-dispatch-history" class="nav-btn whitespace-nowrap md:w-full flex items-center gap-3 px-4 py-3 rounded-xl text-slate-400 hover:bg-slate-800 hover:text-white font-medium transition-colors text-left shrink-0">
                        <i class="ph ph-clock-counter-clockwise text-xl"></i>
                        <span>Dispatch History</span>
                    </button>
                    <button type="button" onclick="switchTab('approvals')" id="nav-approvals" class="nav-btn hidden whitespace-nowrap md:w-full flex items-center gap-3 px-4 py-3 rounded-xl text-slate-400 hover:bg-slate-800 hover:text-white font-medium transition-colors text-left shrink-0">
                        <i class="ph ph-check-square text-xl"></i>
                        <span>Approvals</span>
                    </button>
                    <button type="button" onclick="switchTab('user-management')" id="nav-user-management" class="nav-btn hidden whitespace-nowrap md:w-full flex items-center gap-3 px-4 py-3 rounded-xl text-slate-400 hover:bg-slate-800 hover:text-white font-medium transition-colors text-left shrink-0">
                        <i class="ph ph-users-three text-xl"></i>
                        <span>User Management</span>
                    </button>
                    <button type="button" onclick="requestManageModelsAccess()" id="nav-manage-models" class="hidden whitespace-nowrap md:w-full flex items-center gap-3 px-4 py-3 rounded-xl text-indigo-300 hover:bg-slate-800 hover:text-white font-medium transition-colors text-left shrink-0">
                        <i class="ph ph-gear-six text-xl"></i>
                        <span>Manage Models</span>
                    </button>
                </nav>
                
                <div class="p-4 border-t border-slate-800 hidden md:block shrink-0">
                    <div class="flex items-center justify-between gap-3">
                        <div class="flex items-center gap-3">
                            <div class="w-8 h-8 rounded-full bg-slate-700 flex items-center justify-center">
                                <i class="ph ph-user text-slate-300"></i>
                            </div>
                            <div>
                                <p class="text-xs text-slate-400">Logged in as</p>
                                <p id="user-role-display" class="text-sm font-semibold text-emerald-400">Admin</p>
                            </div>
                        </div>
                        <button type="button" onclick="logout()" class="p-2 text-slate-400 hover:text-rose-400 hover:bg-slate-800 rounded-lg transition" title="Logout">
                            <i class="ph ph-sign-out text-lg"></i>
                        </button>
                    </div>
                </div>
            </aside>

            <!-- Main Content Area -->
            <main class="flex-1 flex flex-col h-full overflow-hidden relative bg-slate-50">
                
                <!-- Header -->
                <header class="h-16 bg-white border-b border-slate-200 flex items-center justify-between px-4 md:px-8 z-10 shrink-0">
                    <div class="flex items-center gap-3">
                        <h2 class="text-xl font-bold text-slate-800" id="header-title">Inventory Dashboard</h2>
                        <button type="button" onclick="openDriveSyncModal()" id="drive-status-badge" class="px-2.5 py-1 rounded-full text-[11px] font-bold border flex items-center gap-1.5 transition bg-slate-100 text-slate-600 border-slate-200" title="Click to configure Google Drive Sync">
                            <span id="drive-status-dot" class="w-2 h-2 rounded-full bg-slate-400"></span>
                            <span id="drive-status-text">Drive: Not Linked</span>
                        </button>
                    </div>
                    <div class="flex items-center gap-2 flex-wrap">
                        <button type="button" id="admin-manage-models-btn" onclick="requestManageModelsAccess()" class="hidden px-3 py-1.5 bg-indigo-50 text-indigo-700 rounded-full text-xs font-semibold border border-indigo-200 items-center gap-1 hover:bg-indigo-100 transition">
                            <i class="ph ph-gear-six"></i> <span>Manage Models</span>
                        </button>
                        <button type="button" onclick="openCsvModal()" class="px-3 py-1.5 bg-emerald-50 text-emerald-700 rounded-full text-xs font-semibold border border-emerald-200 flex items-center gap-1 hover:bg-emerald-100 transition" title="Import CSV">
                            <i class="ph ph-file-arrow-up"></i> <span class="hidden sm:inline">Import CSV</span>
                        </button>
                        <button type="button" onclick="exportToExcel()" class="px-3 py-1.5 bg-teal-50 text-teal-700 rounded-full text-xs font-semibold border border-teal-200 flex items-center gap-1 hover:bg-teal-100 transition" title="Download Excel (.xlsx)">
                            <i class="ph ph-file-xls"></i> <span class="hidden sm:inline">Excel Backup</span>
                        </button>
                        <button type="button" onclick="openDriveSyncModal()" class="px-3 py-1.5 bg-amber-50 text-amber-700 rounded-full text-xs font-semibold border border-amber-200 flex items-center gap-1 hover:bg-amber-100 transition" title="Google Drive Auto-Sync">
                            <i class="ph ph-cloud-arrow-up"></i> <span class="hidden sm:inline">Drive Sync</span>
                        </button>
                        <button type="button" onclick="logout()" class="md:hidden p-2 text-slate-500 hover:text-rose-500 rounded-lg transition" title="Logout">
                            <i class="ph ph-sign-out text-xl"></i>
                        </button>
                    </div>
                </header>

                <!-- Scrollable Viewport Area (Added pb-24 to ensure bottom isn't covered) -->
                <div id="main-scroll-area" class="flex-1 overflow-y-auto overflow-x-hidden p-4 md:p-8 relative pb-24">
                    
                    <!-- Tab: Dashboard -->
                    <div id="tab-dashboard" class="tab-content max-w-6xl mx-auto block space-y-6">
                        <div class="grid grid-cols-1 sm:grid-cols-3 gap-4 md:gap-6">
                            <div class="bg-white p-6 rounded-2xl shadow-sm border border-slate-200">
                                <div class="flex justify-between items-start">
                                    <div>
                                        <p class="text-slate-500 text-sm font-medium mb-1">Total Assets</p>
                                        <h3 class="text-3xl font-bold text-slate-900" id="stat-total">0</h3>
                                    </div>
                                    <div class="w-10 h-10 rounded-full bg-blue-50 flex items-center justify-center text-blue-600"><i class="ph ph-files text-xl"></i></div>
                                </div>
                            </div>
                            <div class="bg-white p-6 rounded-2xl shadow-sm border border-slate-200">
                                <div class="flex justify-between items-start">
                                    <div>
                                        <p class="text-slate-500 text-sm font-medium mb-1">Currently In Stock</p>
                                        <h3 class="text-3xl font-bold text-emerald-600" id="stat-in">0</h3>
                                    </div>
                                    <div class="w-10 h-10 rounded-full bg-emerald-50 flex items-center justify-center text-emerald-600"><i class="ph ph-package text-xl"></i></div>
                                </div>
                            </div>
                            <div class="bg-white p-6 rounded-2xl shadow-sm border border-slate-200">
                                <div class="flex justify-between items-start">
                                    <div>
                                        <p class="text-slate-500 text-sm font-medium mb-1">Dispatched (Out)</p>
                                        <h3 class="text-3xl font-bold text-amber-600" id="stat-out">0</h3>
                                    </div>
                                    <div class="w-10 h-10 rounded-full bg-amber-50 flex items-center justify-center text-amber-600"><i class="ph ph-truck text-xl"></i></div>
                                </div>
                            </div>
                        </div>

                        <!-- Category Breakdown Section -->
                        <div class="bg-white p-6 rounded-2xl shadow-sm border border-slate-200">
                            <div class="flex items-center justify-between mb-4">
                                <h3 class="font-bold text-slate-800 text-base flex items-center gap-2">
                                    <i class="ph ph-chart-bar text-blue-600 text-xl"></i> Category Stock Summary
                                </h3>
                                <span class="text-xs text-slate-500">Live In-Stock Counts & Model Breakdown</span>
                            </div>
                            <div id="category-summary-grid" class="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-4 gap-4">
                                <!-- Populated dynamically -->
                            </div>
                        </div>

                        <!-- Low Stock Dashboard Widget -->
                        <div id="low-stock-widget-container" class="hidden bg-rose-50 border border-rose-200 rounded-2xl p-5 shadow-sm">
                            <div class="flex items-center gap-3 mb-3">
                                <div class="w-9 h-9 rounded-xl bg-rose-600 text-white flex items-center justify-center shrink-0">
                                    <i class="ph-fill ph-warning text-xl"></i>
                                </div>
                                <div>
                                    <h3 class="font-bold text-rose-900 text-base">Low Stock Warnings (Threshold: &le; 5 Units)</h3>
                                    <p class="text-xs text-rose-700">The following models require immediate restocking.</p>
                                </div>
                            </div>
                            <div id="low-stock-list" class="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-3">
                                <!-- Populated dynamically -->
                            </div>
                        </div>

                        <div class="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden flex flex-col">
                            <div class="px-6 py-4 border-b border-slate-200 bg-slate-50 flex flex-col sm:flex-row justify-between items-center gap-4">
                                <h3 class="font-bold text-slate-800">Master Inventory List</h3>
                                <div class="relative w-full sm:w-64">
                                    <i class="ph ph-magnifying-glass absolute left-3 top-1/2 transform -translate-y-1/2 text-slate-400"></i>
                                    <input type="text" id="search-input" placeholder="Search item, source or S/N..." class="w-full pl-9 pr-4 py-2 text-sm border border-slate-300 rounded-lg focus:outline-none focus:border-blue-500 focus:ring-1 focus:ring-blue-500" onkeyup="renderTable()">
                                </div>
                            </div>
                            <div class="overflow-x-auto">
                                <table class="w-full text-left border-collapse min-w-[900px]">
                                    <thead>
                                        <tr class="bg-white text-slate-500 text-xs uppercase tracking-wider border-b border-slate-200">
                                            <th class="px-6 py-4 font-semibold">Category & Model</th>
                                            <th class="px-6 py-4 font-semibold">Serial Number</th>
                                            <th class="px-6 py-4 font-semibold">Status</th>
                                            <th class="px-6 py-4 font-semibold">Activity Info</th>
                                            <th class="px-6 py-4 font-semibold text-right">Actions</th>
                                        </tr>
                                    </thead>
                                    <tbody id="inventory-table-body" class="divide-y divide-slate-100 bg-white">
                                        <!-- Dynamic rows -->
                                    </tbody>
                                </table>
                            </div>
                        </div>
                    </div>

                    <!-- Tab: Material In -->
                    <div id="tab-material-in" class="tab-content max-w-2xl mx-auto hidden">
                        <div class="bg-white rounded-2xl shadow-sm border border-slate-200 p-6 md:p-8">
                            <div class="flex items-center justify-between mb-6">
                                <div class="flex items-center gap-3">
                                    <div class="w-12 h-12 bg-blue-50 text-blue-600 rounded-xl flex items-center justify-center">
                                        <i class="ph ph-box-arrow-down text-2xl"></i>
                                    </div>
                                    <div>
                                        <h3 class="text-xl font-bold text-slate-800">Receive Material</h3>
                                        <p class="text-sm text-slate-500">Log new equipment into the local database.</p>
                                    </div>
                                </div>
                                <button type="button" id="in-tab-manage-btn" onclick="requestManageModelsAccess()" class="hidden px-3 py-2 bg-indigo-50 hover:bg-indigo-100 text-indigo-700 rounded-xl text-xs font-bold transition items-center gap-1.5 border border-indigo-200">
                                    <i class="ph ph-plus-circle text-base"></i> Add / Edit Models
                                </button>
                            </div>

                            <form id="form-material-in" class="space-y-4">
                                <div>
                                    <label class="block text-sm font-semibold text-slate-700 mb-1">Equipment Model / Name <span class="text-rose-500">*</span></label>
                                    <select id="in-name" required class="w-full px-4 py-3 border border-slate-300 rounded-xl focus:ring-2 focus:ring-blue-500 outline-none transition bg-white">
                                        <!-- Populated dynamically via JS -->
                                    </select>
                                </div>
                                
                                <div class="grid grid-cols-1 sm:grid-cols-2 gap-4">
                                    <div>
                                        <label class="block text-sm font-semibold text-slate-700 mb-1">Serial Number (S/N) <span class="text-rose-500">*</span></label>
                                        <input type="text" id="in-sn" required placeholder="e.g. SN-104958" 
                                            class="w-full px-4 py-3 border border-slate-300 rounded-xl focus:ring-2 focus:ring-blue-500 outline-none transition uppercase">
                                    </div>
                                    <div>
                                        <label class="block text-sm font-semibold text-slate-700 mb-1">Received From (Source) <span class="text-rose-500">*</span></label>
                                        <input type="text" id="in-source" required placeholder="e.g. Factory / Vendor / Head Office" 
                                            class="w-full px-4 py-3 border border-slate-300 rounded-xl focus:ring-2 focus:ring-blue-500 outline-none transition">
                                    </div>
                                </div>

                                <div class="grid grid-cols-1 sm:grid-cols-2 gap-4">
                                    <div>
                                        <label class="block text-sm font-semibold text-slate-700 mb-1">Date Received <span class="text-rose-500">*</span></label>
                                        <input type="date" id="in-date" required 
                                            class="w-full px-4 py-3 border border-slate-300 rounded-xl focus:ring-2 focus:ring-blue-500 outline-none transition">
                                    </div>
                                    <div>
                                        <label class="block text-sm font-semibold text-slate-700 mb-1">Accessories Included</label>
                                        <input type="text" id="in-accessories" placeholder="e.g. 2 Batteries, Charger, Hard Case" 
                                            class="w-full px-4 py-3 border border-slate-300 rounded-xl focus:ring-2 focus:ring-blue-500 outline-none transition">
                                    </div>
                                </div>

                                <button type="submit" class="w-full mt-4 bg-blue-600 text-white font-semibold py-3.5 px-4 rounded-xl hover:bg-blue-700 transition duration-200 flex justify-center items-center gap-2 shadow-sm">
                                    <i class="ph ph-plus-circle text-lg"></i> Add to Inventory
                                </button>
                            </form>
                        </div>
                    </div>

                    <!-- Tab: Material Out -->
                    <div id="tab-material-out" class="tab-content max-w-2xl mx-auto hidden">
                        <div class="bg-white rounded-2xl shadow-sm border border-slate-200 p-6 md:p-8">
                            <div class="flex items-center gap-3 mb-6">
                                <div class="w-12 h-12 bg-amber-50 text-amber-600 rounded-xl flex items-center justify-center">
                                    <i class="ph ph-truck text-2xl"></i>
                                </div>
                                <div>
                                    <h3 class="text-xl font-bold text-slate-800">Dispatch Material</h3>
                                    <p class="text-sm text-slate-500">Provide recipient details and choose dispatch mode.</p>
                                </div>
                            </div>

                            <form id="form-material-out" class="space-y-4">
                                <div>
                                    <label class="block text-sm font-semibold text-slate-700 mb-1">Select Equipment (Currently In Stock) <span class="text-rose-500">*</span></label>
                                    <select id="out-item" required class="w-full px-4 py-3 border border-slate-300 rounded-xl focus:ring-2 focus:ring-blue-500 outline-none transition bg-white text-slate-700">
                                        <option value="" disabled selected>Select an item...</option>
                                    </select>
                                </div>

                                <!-- Dispatch Mode Selection -->
                                <div>
                                    <label class="block text-sm font-semibold text-slate-700 mb-2">Dispatch Mode <span class="text-rose-500">*</span></label>
                                    <div class="grid grid-cols-2 gap-4">
                                        <label class="flex items-center gap-3 p-3 border border-slate-300 rounded-xl cursor-pointer hover:bg-slate-50 transition">
                                            <input type="radio" name="dispatch-mode" value="Direct" checked onchange="toggleDispatchMode()" class="text-blue-600 focus:ring-blue-500">
                                            <div>
                                                <span class="font-semibold text-sm block text-slate-800">Direct / Hand Delivery</span>
                                                <span class="text-xs text-slate-500">Standard site challan</span>
                                            </div>
                                        </label>
                                        <label class="flex items-center gap-3 p-3 border border-slate-300 rounded-xl cursor-pointer hover:bg-slate-50 transition">
                                            <input type="radio" name="dispatch-mode" value="Courier" onchange="toggleDispatchMode()" class="text-blue-600 focus:ring-blue-500">
                                            <div>
                                                <span class="font-semibold text-sm block text-slate-800">Courier / Logistics</span>
                                                <span class="text-xs text-slate-500">Shipped via service</span>
                                            </div>
                                        </label>
                                    </div>
                                </div>
                                
                                <!-- Courier Specific Fields (Hidden by default) -->
                                <div id="courier-fields" class="hidden space-y-4 p-4 bg-slate-50 rounded-xl border border-slate-200">
                                    <h4 class="text-xs font-bold uppercase tracking-wider text-slate-500 flex items-center gap-1.5">
                                        <i class="ph ph-package text-blue-600"></i> Courier Logistics Details
                                    </h4>
                                    <div class="grid grid-cols-1 sm:grid-cols-2 gap-4">
                                        <div>
                                            <label class="block text-sm font-semibold text-slate-700 mb-1">Courier Partner Name <span class="text-rose-500">*</span></label>
                                            <input type="text" id="out-courier-name" placeholder="e.g. Blue Dart / DTDC / Professional" 
                                                class="w-full px-4 py-2.5 border border-slate-300 rounded-xl focus:ring-2 focus:ring-blue-500 outline-none transition bg-white text-sm">
                                        </div>
                                        <div>
                                            <label class="block text-sm font-semibold text-slate-700 mb-1">AWB / Tracking Number <span class="text-slate-400 font-normal">(Optional)</span></label>
                                            <input type="text" id="out-awb" placeholder="e.g. AW-984729104" 
                                                class="w-full px-4 py-2.5 border border-slate-300 rounded-xl focus:ring-2 focus:ring-blue-500 outline-none transition bg-white uppercase text-sm">
                                        </div>
                                    </div>
                                </div>
                                
                                <div class="grid grid-cols-1 sm:grid-cols-2 gap-4">
                                    <div>
                                        <label class="block text-sm font-semibold text-slate-700 mb-1">Company / Client Name <span class="text-rose-500">*</span></label>
                                        <input type="text" id="out-company" required placeholder="e.g. Acme Surveyors Pvt Ltd" 
                                            class="w-full px-4 py-3 border border-slate-300 rounded-xl focus:ring-2 focus:ring-blue-500 outline-none transition">
                                    </div>
                                    <div>
                                        <label class="block text-sm font-semibold text-slate-700 mb-1">Challan / Invoice Number <span class="text-slate-400 font-normal">(Optional)</span></label>
                                        <input type="text" id="out-challan" placeholder="e.g. INV-2026-089" 
                                            class="w-full px-4 py-3 border border-slate-300 rounded-xl focus:ring-2 focus:ring-blue-500 outline-none transition">
                                    </div>
                                </div>

                                <div class="grid grid-cols-1 sm:grid-cols-2 gap-4">
                                    <div>
                                        <label class="block text-sm font-semibold text-slate-700 mb-1">Contact Phone Number <span class="text-rose-500">*</span></label>
                                        <input type="text" id="out-phone" required placeholder="e.g. +91 9876543210" 
                                            class="w-full px-4 py-3 border border-slate-300 rounded-xl focus:ring-2 focus:ring-blue-500 outline-none transition">
                                    </div>
                                    <div>
                                        <label class="block text-sm font-semibold text-slate-700 mb-1">Site / Delivery Location <span class="text-rose-500">*</span></label>
                                        <input type="text" id="out-to" required placeholder="e.g. Sector 62, Noida Site A" 
                                            class="w-full px-4 py-3 border border-slate-300 rounded-xl focus:ring-2 focus:ring-blue-500 outline-none transition">
                                    </div>
                                </div>

                                <div>
                                    <label class="block text-sm font-semibold text-slate-700 mb-1">Full Details / Remarks / Accessories</label>
                                    <textarea id="out-remarks" rows="2" placeholder="e.g. Issued with 2 batteries, charger, hard case, calibration certificate." 
                                        class="w-full px-4 py-3 border border-slate-300 rounded-xl focus:ring-2 focus:ring-blue-500 outline-none transition"></textarea>
                                </div>

                                <button type="submit" class="w-full mt-4 bg-slate-900 text-white font-semibold py-3.5 px-4 rounded-xl hover:bg-slate-800 transition duration-200 flex justify-center items-center gap-2 shadow-sm">
                                    <i class="ph ph-printer text-lg"></i> Request Dispatch
                                </button>
                            </form>
                        </div>
                    </div>

                    <!-- Tab: Dispatch History -->
                    <div id="tab-dispatch-history" class="tab-content max-w-6xl mx-auto hidden">
                        <div class="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden flex flex-col">
                            <div class="px-6 py-4 border-b border-slate-200 bg-slate-50 flex flex-col md:flex-row justify-between items-center gap-4">
                                <div>
                                    <h3 class="font-bold text-slate-800 text-lg">Dispatch History Log</h3>
                                    <p class="text-xs text-slate-500">Filter and audit all past dispatched equipment date-wise.</p>
                                </div>
                                <div class="flex flex-wrap items-center gap-3 w-full md:w-auto">
                                    <div class="flex items-center gap-2 text-xs">
                                        <span class="text-slate-500 font-medium">From:</span>
                                        <input type="date" id="filter-date-from" onchange="renderDispatchHistory()" class="px-3 py-1.5 border border-slate-300 rounded-lg text-xs bg-white outline-none focus:border-blue-500">
                                    </div>
                                    <div class="flex items-center gap-2 text-xs">
                                        <span class="text-slate-500 font-medium">To:</span>
                                        <input type="date" id="filter-date-to" onchange="renderDispatchHistory()" class="px-3 py-1.5 border border-slate-300 rounded-lg text-xs bg-white outline-none focus:border-blue-500">
                                    </div>
                                    <button type="button" onclick="resetDispatchFilters()" class="px-3 py-1.5 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-lg text-xs font-semibold transition">Reset</button>
                                </div>
                            </div>
                            <div class="overflow-x-auto">
                                <table class="w-full text-left border-collapse min-w-[900px]">
                                    <thead>
                                        <tr class="bg-white text-slate-500 text-xs uppercase tracking-wider border-b border-slate-200">
                                            <th class="px-6 py-4 font-semibold">Equipment & S/N</th>
                                            <th class="px-6 py-4 font-semibold">Client & Challan</th>
                                            <th class="px-6 py-4 font-semibold">Dispatch Mode & Info</th>
                                            <th class="px-6 py-4 font-semibold">Date Out</th>
                                            <th class="px-6 py-4 font-semibold text-right">Actions</th>
                                        </tr>
                                    </thead>
                                    <tbody id="dispatch-history-table-body" class="divide-y divide-slate-100 bg-white">
                                        <!-- Dynamic rows -->
                                    </tbody>
                                </table>
                            </div>
                        </div>
                    </div>

                    <!-- Tab: Approvals -->
                    <div id="tab-approvals" class="tab-content max-w-4xl mx-auto hidden">
                        <div class="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden flex flex-col">
                            <div class="px-6 py-4 border-b border-slate-200 bg-slate-50 flex items-center gap-3">
                                <i class="ph-fill ph-shield-check text-blue-600 text-xl"></i>
                                <h3 class="font-bold text-slate-800">Pending Dispatch Approvals</h3>
                            </div>
                            <div class="overflow-x-auto">
                                <table class="w-full text-left border-collapse">
                                    <thead>
                                        <tr class="bg-white text-slate-500 text-xs uppercase tracking-wider border-b border-slate-200">
                                            <th class="px-6 py-4 font-semibold">Equipment & S/N</th>
                                            <th class="px-6 py-4 font-semibold">Client & Mode</th>
                                            <th class="px-6 py-4 font-semibold text-right">Actions</th>
                                        </tr>
                                    </thead>
                                    <tbody id="approvals-table-body" class="divide-y divide-slate-100 bg-white">
                                        <!-- Dynamic rows -->
                                    </tbody>
                                </table>
                            </div>
                        </div>
                    </div>

                    <!-- Tab: User Management -->
                    <div id="tab-user-management" class="tab-content max-w-5xl mx-auto hidden space-y-6">
                        <div class="bg-white rounded-2xl shadow-sm border border-slate-200 p-6 md:p-8">
                            <div class="flex items-center gap-3 mb-6">
                                <div class="w-12 h-12 bg-indigo-50 text-indigo-600 rounded-xl flex items-center justify-center">
                                    <i class="ph ph-user-plus text-2xl"></i>
                                </div>
                                <div>
                                    <h3 class="text-xl font-bold text-slate-800">Create New User Account</h3>
                                    <p class="text-sm text-slate-500">Add an administrator or store manager credential.</p>
                                </div>
                            </div>

                            <form id="form-create-user" onsubmit="handleCreateUser(event)" class="grid grid-cols-1 sm:grid-cols-3 gap-4">
                                <div>
                                    <label class="block text-sm font-semibold text-slate-700 mb-1">Username <span class="text-rose-500">*</span></label>
                                    <input type="text" id="new-user-name" required placeholder="e.g. john_manager" class="w-full px-4 py-2.5 border border-slate-300 rounded-xl text-sm outline-none focus:ring-2 focus:ring-blue-500 bg-white">
                                </div>
                                <div>
                                    <label class="block text-sm font-semibold text-slate-700 mb-1">Password <span class="text-rose-500">*</span></label>
                                    <input type="password" id="new-user-pass" required placeholder="••••••••" class="w-full px-4 py-2.5 border border-slate-300 rounded-xl text-sm outline-none focus:ring-2 focus:ring-blue-500 bg-white">
                                </div>
                                <div>
                                    <label class="block text-sm font-semibold text-slate-700 mb-1">Role <span class="text-rose-500">*</span></label>
                                    <div class="flex gap-2">
                                        <select id="new-user-role" required class="flex-1 px-4 py-2.5 border border-slate-300 rounded-xl text-sm outline-none focus:ring-2 focus:ring-blue-500 bg-white">
                                            <option value="manager">Store Manager</option>
                                            <option value="admin">Admin</option>
                                        </select>
                                        <button type="submit" class="px-5 py-2.5 bg-indigo-600 hover:bg-indigo-700 text-white font-semibold rounded-xl text-sm transition shadow-sm shrink-0">
                                            Create User
                                        </button>
                                    </div>
                                </div>
                            </form>
                        </div>

                        <!-- User Accounts Table -->
                        <div class="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden flex flex-col">
                            <div class="px-6 py-4 border-b border-slate-200 bg-slate-50 flex items-center justify-between">
                                <div>
                                    <h3 class="font-bold text-slate-800 text-base">Registered System Accounts</h3>
                                    <p class="text-xs text-slate-500">Manage usernames, roles, or remove accounts.</p>
                                </div>
                                <span id="user-count-badge" class="px-2.5 py-1 bg-slate-200 text-slate-700 rounded-full text-xs font-bold">0 Users</span>
                            </div>
                            <div class="overflow-x-auto">
                                <table class="w-full text-left border-collapse">
                                    <thead>
                                        <tr class="bg-white text-slate-500 text-xs uppercase tracking-wider border-b border-slate-200">
                                            <th class="px-6 py-4 font-semibold">Username</th>
                                            <th class="px-6 py-4 font-semibold">Role</th>
                                            <th class="px-6 py-4 font-semibold">Created Date</th>
                                            <th class="px-6 py-4 font-semibold text-right">Actions</th>
                                        </tr>
                                    </thead>
                                    <tbody id="users-table-body" class="divide-y divide-slate-100 bg-white">
                                        <!-- Dynamic rows -->
                                    </tbody>
                                </table>
                            </div>
                        </div>
                    </div>

                </div>
            </main>
        </div>

        <!-- Password Prompt Modal for Manage Models (Passcode: 8317) -->
        <div id="model-pin-modal" class="hidden fixed inset-0 z-[110] bg-slate-900/70 backdrop-blur-sm flex items-center justify-center p-4">
            <div class="bg-white p-6 rounded-2xl shadow-2xl max-w-xs w-full text-center space-y-4">
                <div class="w-12 h-12 bg-indigo-50 text-indigo-600 rounded-xl flex items-center justify-center mx-auto text-2xl">
                    <i class="ph-fill ph-lock-key"></i>
                </div>
                <div>
                    <h3 class="font-bold text-slate-800 text-lg">Enter Access PIN</h3>
                    <p class="text-xs text-slate-500 mt-0.5">Please enter PIN to manage models (8317)</p>
                </div>
                <div id="pin-error" class="hidden text-rose-600 text-xs font-semibold">Incorrect PIN. Try again (8317).</div>
                <form onsubmit="verifyModelPin(event)" class="space-y-3">
                    <input type="password" id="model-pin-input" autofocus maxlength="4" placeholder="••••" required class="w-full px-4 py-3 text-center text-xl tracking-widest border border-slate-300 rounded-xl focus:ring-2 focus:ring-indigo-500 outline-none font-mono">
                    <div class="flex gap-2">
                        <button type="button" onclick="closeModelPinModal()" class="flex-1 py-2.5 bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-semibold rounded-xl transition">Cancel</button>
                        <button type="submit" class="flex-1 py-2.5 bg
