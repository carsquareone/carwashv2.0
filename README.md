<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0, user-scalable=yes">
    <title>8068 CAR SQUARE | POS System v7.4</title>
    <!-- Firebase SDKs -->
    <script src="https://www.gstatic.com/firebasejs/9.22.0/firebase-app-compat.js"></script>
    <script src="https://www.gstatic.com/firebasejs/9.22.0/firebase-firestore-compat.js"></script>
    <script src="https://www.gstatic.com/firebasejs/9.22.0/firebase-auth-compat.js"></script>
    <script src="https://cdn.tailwindcss.com"></script>
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap" rel="stylesheet">
    <script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
    <style>
        * { font-family: 'Inter', sans-serif; }
        body { background: #0a0a0a; overflow-x: hidden; }
        ::-webkit-scrollbar { width: 8px; }
        ::-webkit-scrollbar-track { background: #1a1a1a; }
        ::-webkit-scrollbar-thumb { background: #00d4ff; border-radius: 10px; }
        .neon-glow { box-shadow: 0 0 12px rgba(0, 212, 255, 0.6); }
        .neon-border { border: 1px solid #00d4ff; box-shadow: 0 0 8px rgba(0,212,255,0.3); }
        .btn-neon { background: linear-gradient(135deg, #00b4d8, #0077b6); transition: all 0.2s; }
        .btn-neon:hover { transform: scale(1.02); box-shadow: 0 0 15px #00d4ff; }
        .card-premium { background: rgba(10, 10, 10, 0.9); backdrop-filter: blur(2px); border: 1px solid rgba(0, 212, 255, 0.3); border-radius: 1rem; }
        .service-btn { background: #1a1a1a; border: 1px solid #333; transition: all 0.2s; cursor: pointer; }
        .service-btn:hover { border-color: #00d4ff; transform: translateY(-2px); background: #222; }
        .glow-text { text-shadow: 0 0 5px #00d4ff; }
        .modal { position: fixed; inset: 0; background: rgba(0,0,0,0.95); display: flex; align-items: center; justify-content: center; z-index: 1000; }
        .modal-content { background: #1a1a1a; border-radius: 1rem; padding: 1.5rem; max-width: 800px; width: 90%; border: 1px solid #00d4ff; max-height: 85vh; overflow-y: auto; }
        .commission-badge { background: linear-gradient(135deg, #00d4ff20, #0077b620); border-radius: 20px; padding: 4px 12px; font-size: 12px; font-weight: bold; }
        .editable-field { cursor: pointer; transition: all 0.2s; }
        .editable-field:hover { background: rgba(0,212,255,0.15); border-radius: 8px; padding: 2px 6px; }
        .admin-grid { display: grid; grid-template-columns: 260px 1fr; gap: 20px; }
        .admin-sidebar { background: rgba(10,10,10,0.95); border-right: 1px solid #00d4ff; border-radius: 1rem 0 0 1rem; padding: 20px; }
        .admin-sidebar-item { padding: 10px 14px; margin: 4px 0; border-radius: 10px; cursor: pointer; transition: all 0.2s; font-size: 14px; }
        .admin-sidebar-item:hover { background: rgba(0,212,255,0.1); }
        .admin-sidebar-item.active { background: rgba(0,212,255,0.2); border-left: 3px solid #00d4ff; }
        .admin-content { padding: 20px; max-height: 80vh; overflow-y: auto; }
        .action-btn { padding: 4px 10px; border-radius: 6px; font-size: 12px; cursor: pointer; transition: all 0.2s; margin: 0 2px; }
        .action-btn:hover { transform: scale(1.05); }
        .excel-export-btn { background: linear-gradient(135deg, #217346, #1a6e3a); transition: all 0.2s; }
        .excel-export-btn:hover { transform: scale(1.05); box-shadow: 0 0 20px rgba(33,115,70,0.5); }
        .open-service-card { transition: all 0.3s; }
        .open-service-card:hover { transform: translateY(-3px); box-shadow: 0 8px 30px rgba(0,212,255,0.1); }
        .status-badge { padding: 2px 12px; border-radius: 12px; font-size: 10px; font-weight: bold; }
        .status-pending { background: #ffaa0033; color: #ffaa00; }
        .status-in-progress { background: #00d4ff33; color: #00d4ff; }
        .status-completed { background: #00ff8833; color: #00ff88; }
        .commission-item { animation: slideIn 0.3s ease; }
        @keyframes slideIn {
            from { opacity: 0; transform: translateX(-10px); }
            to { opacity: 1; transform: translateX(0); }
        }
        .live-dot { display: inline-block; width: 8px; height: 8px; background: #00ff88; border-radius: 50%; animation: pulse 1.5s infinite; }
        @keyframes pulse {
            0% { opacity: 1; transform: scale(1); }
            50% { opacity: 0.5; transform: scale(0.8); }
            100% { opacity: 1; transform: scale(1); }
        }
        .sync-status { display: inline-block; padding: 4px 12px; border-radius: 20px; font-size: 11px; font-weight: bold; }
        .sync-active { background: #00ff8833; color: #00ff88; }
        .sync-inactive { background: #ff6b6b33; color: #ff6b6b; }
        .sync-syncing { background: #ffaa0033; color: #ffaa00; animation: pulse 1s infinite; }
        .commission-detail-row { border-left: 3px solid #00ff8833; }
        .firebase-status { display: inline-block; padding: 4px 12px; border-radius: 20px; font-size: 11px; font-weight: bold; }
        .firebase-connected { background: #00ff8833; color: #00ff88; }
        .firebase-disconnected { background: #ff6b6b33; color: #ff6b6b; }
        .firebase-connecting { background: #ffaa0033; color: #ffaa00; animation: pulse 1s infinite; }
        @media print { .no-print { display: none; } .print-area { display: block; } }
        @media (max-width: 768px) { .admin-grid { grid-template-columns: 1fr; } .admin-sidebar { border-radius: 1rem; } }
    </style>
</head>
<body class="text-white">

<div id="root" class="min-h-screen bg-black"></div>

<script>
    // ======================== FIREBASE CONFIGURATION ========================
    // Replace with your own Firebase config
    const firebaseConfig = {
        apiKey: "YOUR_API_KEY",
        authDomain: "YOUR_PROJECT.firebaseapp.com",
        projectId: "YOUR_PROJECT_ID",
        storageBucket: "YOUR_PROJECT.appspot.com",
        messagingSenderId: "YOUR_SENDER_ID",
        appId: "YOUR_APP_ID"
    };

    // Initialize Firebase
    let firebaseApp = null;
    let firestore = null;
    let auth = null;
    let firebaseInitialized = false;

    function initFirebase() {
        try {
            if (!firebaseConfig.apiKey || firebaseConfig.apiKey === "YOUR_API_KEY") {
                console.warn("Firebase not configured. Using localStorage mode.");
                return false;
            }
            
            firebaseApp = firebase.initializeApp(firebaseConfig);
            firestore = firebase.firestore(firebaseApp);
            auth = firebase.auth(firebaseApp);
            firebaseInitialized = true;
            
            // Enable offline persistence
            firestore.enablePersistence({ synchronizeTabs: true })
                .then(() => console.log("Firebase offline persistence enabled"))
                .catch(err => console.warn("Firebase persistence error:", err));
            
            console.log("Firebase initialized successfully!");
            return true;
        } catch(e) {
            console.error("Firebase initialization error:", e);
            return false;
        }
    }

    // ======================== FIREBASE DATA SYNC ========================
    async function syncToFirebase(collection, data) {
        if (!firebaseInitialized) return false;
        try {
            const docRef = firestore.collection(collection).doc('main');
            await docRef.set({
                data: data,
                updatedAt: firebase.firestore.FieldValue.serverTimestamp()
            });
            return true;
        } catch(e) {
            console.error("Firebase sync error:", e);
            return false;
        }
    }

    async function loadFromFirebase(collection) {
        if (!firebaseInitialized) return null;
        try {
            const docRef = firestore.collection(collection).doc('main');
            const doc = await docRef.get();
            if (doc.exists) {
                return doc.data().data;
            }
            return null;
        } catch(e) {
            console.error("Firebase load error:", e);
            return null;
        }
    }

    async function syncAllToFirebase() {
        if (!firebaseInitialized) {
            updateFirebaseStatus('disconnected');
            return false;
        }
        
        updateFirebaseStatus('connecting');
        
        try {
            const dataToSync = {
                services: appData.services,
                categories: appData.categories,
                customers: appData.customers,
                transactions: appData.transactions,
                openServices: appData.openServices,
                staff: appData.staff,
                expenses: appData.expenses,
                inventory: appData.inventory,
                systemSettings: appData.systemSettings,
                commissionMatrix: appData.commissionMatrix,
                nextIds: appData.nextIds
            };
            
            const result = await syncToFirebase('pos_system', dataToSync);
            if (result) {
                appData.systemSettings.lastSync = new Date().toISOString();
                appData.systemSettings.firebaseSync = true;
                saveToLocal();
                updateFirebaseStatus('connected');
                return true;
            }
            return false;
        } catch(e) {
            console.error("Sync all error:", e);
            updateFirebaseStatus('disconnected');
            return false;
        }
    }

    async function loadAllFromFirebase() {
        if (!firebaseInitialized) {
            updateFirebaseStatus('disconnected');
            return false;
        }
        
        updateFirebaseStatus('connecting');
        
        try {
            const data = await loadFromFirebase('pos_system');
            if (data) {
                // Merge loaded data with existing
                Object.assign(appData, data);
                appData.systemSettings.lastSync = new Date().toISOString();
                appData.systemSettings.firebaseSync = true;
                saveToLocal();
                updateFirebaseStatus('connected');
                renderApp();
                return true;
            }
            return false;
        } catch(e) {
            console.error("Load all error:", e);
            updateFirebaseStatus('disconnected');
            return false;
        }
    }

    function updateFirebaseStatus(status) {
        const statusEl = document.getElementById('firebaseStatus');
        if (statusEl) {
            if (status === 'connected') {
                statusEl.innerHTML = '<i class="fas fa-check-circle"></i> Firebase Connected';
                statusEl.className = 'firebase-status firebase-connected';
            } else if (status === 'connecting') {
                statusEl.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Connecting...';
                statusEl.className = 'firebase-status firebase-connecting';
            } else {
                statusEl.innerHTML = '<i class="fas fa-cloud"></i> Firebase Offline';
                statusEl.className = 'firebase-status firebase-disconnected';
            }
        }
    }

    // ======================== COMMISSION MATRIX ========================
    function initializeCommissionMatrix() {
        return {
            "BASIC CARWASH_Small": [52, 65, 70], "BASIC CARWASH_Medium": [62, 72, 80],
            "BASIC CARWASH_Large": [70, 83, 95], "BASIC CARWASH_XL": [85, 100, 115],
            "BASIC CARWASH_2XL": [100, 117, 132], "BASIC CARWASH_3XL": [115, 135, 152],
            "ARMOR CARWASH_Small": [83, 95, 110], "ARMOR CARWASH_Medium": [92, 108, 125],
            "ARMOR CARWASH_Large": [100, 117, 135], "ARMOR CARWASH_XL": [115, 140, 158],
            "ARMOR CARWASH_2XL": [128, 150, 170], "ARMOR CARWASH_3XL": [144, 168, 192],
            "WAX CARWASH_Small": [150, 180, 200], "WAX CARWASH_Medium": [168, 200, 220],
            "WAX CARWASH_Large": [185, 215, 240], "WAX CARWASH_XL": [198, 230, 265],
            "WAX CARWASH_2XL": [215, 250, 285], "WAX CARWASH_3XL": [230, 265, 310],
            "WAX ARMOR CARWASH_Small": [185, 215, 245], "WAX ARMOR CARWASH_Medium": [200, 230, 270],
            "WAX ARMOR CARWASH_Large": [215, 250, 285], "WAX ARMOR CARWASH_XL": [230, 267, 305],
            "WAX ARMOR CARWASH_2XL": [245, 285, 325], "WAX ARMOR CARWASH_3XL": [260, 300, 340],
            "BUFFING CARWASH_Small": [280, 330, 370], "BUFFING CARWASH_Medium": [290, 340, 390],
            "BUFFING CARWASH_Large": [340, 400, 450], "BUFFING CARWASH_XL": [370, 430, 490],
            "BUFFING CARWASH_2XL": [410, 475, 550], "BACK TO 0 CARWASH_Small": [205, 240, 275],
            "BACK TO 0 CARWASH_Medium": [210, 245, 280], "BACK TO 0 CARWASH_Large": [220, 260, 295],
            "BACK TO 0 CARWASH_XL": [235, 280, 315], "BACK TO 0 CARWASH_2XL": [250, 295, 335],
            "BACK TO 0 CARWASH_3XL": [265, 308, 355], "PLATINUM CARWASH_Small": [300, 350, 400],
            "PLATINUM CARWASH_Medium": [330, 385, 440], "PLATINUM CARWASH_Large": [360, 420, 480],
            "PLATINUM CARWASH_XL": [390, 455, 520], "PLATINUM CARWASH_2XL": [420, 490, 560],
            "PLATINUM CARWASH_3XL": [450, 525, 600], "DIAMOND CARWASH_Small": [690, 805, 920],
            "DIAMOND CARWASH_Medium": [720, 840, 960], "DIAMOND CARWASH_Large": [780, 910, 1040],
            "DIAMOND CARWASH_XL": [810, 945, 1080], "DIAMOND CARWASH_2XL": [840, 980, 1120],
            "DIAMOND CARWASH_3XL": [870, 1015, 1160], "ACID RAIN CARWASH_Small": [300, 350, 400],
            "ACID RAIN CARWASH_Medium": [360, 420, 480], "ACID RAIN CARWASH_Large": [390, 455, 520],
            "ACID RAIN CARWASH_XL": [420, 490, 560], "ACID RAIN CARWASH_2XL": [450, 525, 600],
            "ACID RAIN CARWASH_3XL": [480, 560, 640]
        };
    }

    const defaultCustomerPrices = {
        "BASIC CARWASH": { Small: 200, Medium: 230, Large: 250, XL: 300, "2XL": 350, "3XL": 400 },
        "ARMOR CARWASH": { Small: 300, Medium: 330, Large: 350, XL: 420, "2XL": 450, "3XL": 480 },
        "WAX CARWASH": { Small: 550, Medium: 600, Large: 650, XL: 700, "2XL": 750, "3XL": 800 },
        "WAX ARMOR CARWASH": { Small: 650, Medium: 700, Large: 750, XL: 800, "2XL": 850, "3XL": 900 },
        "BUFFING CARWASH": { Small: 1000, Medium: 1050, Large: 1200, XL: 1300, "2XL": 1400 },
        "BACK TO 0 CARWASH": { Small: 680, Medium: 700, Large: 735, XL: 785, "2XL": 835, "3XL": 885 },
        "PLATINUM CARWASH": { Small: 1000, Medium: 1100, Large: 1200, XL: 1300, "2XL": 1400, "3XL": 1500 },
        "DIAMOND CARWASH": { Small: 2300, Medium: 2400, Large: 2600, XL: 2700, "2XL": 2800, "3XL": 2900 },
        "ACID RAIN CARWASH": { Small: 1000, Medium: 1200, Large: 1300, XL: 1400, "2XL": 1500, "3XL": 1600 },
        "INTERIOR DETAILING": { Standard: 4500 }, "EXTERIOR DETAILING": { Standard: 4500 },
        "TINT INSTALLATION": { Standard: 3800 }, "CERAMIC COATING": { Standard: 8000 },
        "GRAPHENE COATING": { Standard: 10000 }, "MOTOR WASH": { Small: 120, Large: 170 }
    };

    // ======================== APP DATA ========================
    let appData = {
        services: Object.entries(defaultCustomerPrices).map(([name, prices], idx) => ({
            id: idx + 1, name: name, desc: getDescription(name), sizes: Object.keys(prices), prices: prices, category: getCategory(name), isActive: true
        })),
        categories: ['Basic Wash', 'Premium Wash', 'Wax Services', 'Buffing', 'Restoration', 'Platinum', 'Diamond', 'Acid Rain', 'Specialty'],
        customers: [],
        transactions: [],
        openServices: [],
        staff: [],
        expenses: [],
        inventory: [
            { id: 1, name: "Car Shampoo", category: "Supplies", quantity: 50, unit: "Liters", reorderLevel: 20 },
            { id: 2, name: "Microfiber Towels", category: "Equipment", quantity: 200, unit: "Pieces", reorderLevel: 50 },
            { id: 3, name: "Tire Black", category: "Supplies", quantity: 30, unit: "Liters", reorderLevel: 10 },
            { id: 4, name: "Wax Polish", category: "Supplies", quantity: 15, unit: "Bottles", reorderLevel: 5 },
            { id: 5, name: "Interior Cleaner", category: "Supplies", quantity: 25, unit: "Liters", reorderLevel: 8 }
        ],
        systemSettings: {
            companyName: "8068 CAR SQUARE",
            tagline: "One Stop Shop",
            address: "8068 Carsquare, Dr Arcadio Santos Ave, Parañaque, 1720 Metro Manila",
            phone: "+63 947 702 2630 | +63 961 519 5229",
            facebook: "8068 Carsquare PH",
            taxRate: 0,
            loyaltyPointsPerPeso: 0.01,
            receiptFooter: "Drive Clean. Drive Premium.",
            commissionSplitRule: "50/50 Equal Split",
            currencySymbol: "₱",
            timezone: "Asia/Manila",
            businessHours: "Mon-Sat 8:00 AM - 7:00 PM",
            companyLogo: null,
            defaultPaymentMethods: ['Cash', 'GCash', 'Maya', 'Credit Card', 'Bank Transfer'],
            firebaseSync: false,
            lastSync: null,
            syncToken: null
        },
        users: [
            { id: 1, username: "owner", password: "owner123", role: "Owner", name: "Owner", isActive: true },
            { id: 2, username: "manager", password: "manager123", role: "Manager", name: "Manager", isActive: true },
            { id: 3, username: "cashier", password: "cashier123", role: "Cashier", name: "Cashier", isActive: true },
            { id: 4, username: "staff", password: "staff123", role: "Staff", name: "Staff", isActive: true }
        ],
        currentUser: null,
        nextIds: { transaction: 100, customer: 50, expense: 20, service: 50, openService: 1 },
        commissionMatrix: initializeCommissionMatrix()
    };

    function getDescription(name) {
        const desc = {
            "BASIC CARWASH": "Carwash + Vacuum + Tire Black",
            "ARMOR CARWASH": "Carwash + Vacuum + Interior Dressing + Tire Black",
            "WAX CARWASH": "Carwash + Vacuum + Hydrophobic Wax + Tire Black",
            "WAX ARMOR CARWASH": "Carwash + Vacuum + Interior Dressing + Hydrophobic Wax + Tire Black",
            "BUFFING CARWASH": "Carwash + Vacuum + Interior Dressing + Hydrophobic Wax + Buffing Wax + Tire Black",
            "BACK TO 0 CARWASH": "Complete restoration + Paint correction",
            "PLATINUM CARWASH": "Carwash + Vacuum + Interior Dressing + Ceramic Wax + Tire Black",
            "DIAMOND CARWASH": "Carwash + Vacuum + Interior Dressing + Buffing Wax + Ceramic Wax + Tire Black",
            "ACID RAIN CARWASH": "Acid rain removal + Paint decontamination"
        };
        return desc[name] || "Premium automotive service";
    }

    function getCategory(name) {
        if(name.includes("BASIC")) return "Basic Wash";
        if(name.includes("ARMOR") && !name.includes("WAX")) return "Premium Wash";
        if(name.includes("WAX ARMOR")) return "Premium Wax";
        if(name.includes("WAX")) return "Wax Services";
        if(name.includes("BUFFING")) return "Buffing";
        if(name.includes("BACK")) return "Restoration";
        if(name.includes("PLATINUM")) return "Platinum";
        if(name.includes("DIAMOND")) return "Diamond";
        if(name.includes("ACID")) return "Acid Rain";
        return "Specialty";
    }

    // ======================== LOAD SAVED DATA ========================
    const saved = localStorage.getItem("8068_enterprise_v7_4");
    if(saved) {
        try {
            const parsed = JSON.parse(saved);
            appData = { ...appData, ...parsed };
            if (!appData.commissionMatrix || Object.keys(appData.commissionMatrix).length === 0) {
                appData.commissionMatrix = initializeCommissionMatrix();
            }
            if (!appData.openServices) appData.openServices = [];
            if (!appData.inventory) appData.inventory = [];
            if (!appData.users) appData.users = [
                { id: 1, username: "owner", password: "owner123", role: "Owner", name: "Owner", isActive: true },
                { id: 2, username: "manager", password: "manager123", role: "Manager", name: "Manager", isActive: true },
                { id: 3, username: "cashier", password: "cashier123", role: "Cashier", name: "Cashier", isActive: true },
                { id: 4, username: "staff", password: "staff123", role: "Staff", name: "Staff", isActive: true }
            ];
            if (!appData.currentUser) appData.currentUser = appData.users[0];
            if (!appData.systemSettings.defaultPaymentMethods) {
                appData.systemSettings.defaultPaymentMethods = ['Cash', 'GCash', 'Maya', 'Credit Card', 'Bank Transfer'];
            }
        } catch(e) {}
    }

    function saveToLocal() {
        localStorage.setItem("8068_enterprise_v7_4", JSON.stringify(appData));
        localStorage.setItem("8068_system_last_saved", new Date().toISOString());
        
        // Auto-sync to Firebase if enabled
        if (appData.systemSettings.firebaseSync && firebaseInitialized) {
            syncAllToFirebase();
        }
    }

    function getCommissionAmount(serviceName, size, tierPercent) {
        const key = `${serviceName}_${size}`;
        const matrix = appData.commissionMatrix || {};
        const values = matrix[key];
        if (!values) return 0;
        if (tierPercent === 30) return values[0] || 0;
        if (tierPercent === 35) return values[1] || 0;
        if (tierPercent === 40) return values[2] || 0;
        return 0;
    }

    // ======================== STATE ========================
    let currentView = "dashboard";
    let currentAdminPanel = "services";
    let posCart = [];
    let selectedCustomer = null;
    let vehicleInfo = { plate: "", model: "", size: "Medium" };
    let discount = 0;
    let loyaltyPointsUsed = 0;
    let selectedPayment = "Cash";
    let cashTendered = 0;
    let cartRemarks = "";
    let showLogin = true;

    function generateReceiptNo() { return "INV-" + new Date().getFullYear() + "-" + (appData.nextIds.transaction++); }

    // ======================== AUTHENTICATION ========================
    function login(username, password) {
        const user = appData.users.find(u => u.username === username && u.password === password && u.isActive);
        if(user) {
            appData.currentUser = user;
            showLogin = false;
            saveToLocal();
            renderApp();
            return true;
        }
        return false;
    }

    function logout() {
        appData.currentUser = null;
        showLogin = true;
        renderApp();
    }

    function hasRole(roles) {
        if(!appData.currentUser) return false;
        return roles.includes(appData.currentUser.role);
    }

    function renderLogin() {
        return `<div class="min-h-screen flex items-center justify-center">
            <div class="card-premium p-8 max-w-md w-full">
                <div class="text-center mb-6">
                    ${appData.systemSettings.companyLogo ? `<img src="${appData.systemSettings.companyLogo}" class="mx-auto max-h-20 mb-4">` : ''}
                    <h2 class="text-3xl font-bold glow-text">${appData.systemSettings.companyName}</h2>
                    <p class="text-gray-400 text-sm">${appData.systemSettings.tagline}</p>
                    <p class="text-cyan-400 text-xs mt-2">Enterprise POS System v7.4</p>
                </div>
                <div class="space-y-4">
                    <div><label class="text-gray-400">Username</label><input type="text" id="loginUser" class="w-full bg-gray-800 p-3 rounded border border-gray-700" placeholder="Enter username" value="owner"></div>
                    <div><label class="text-gray-400">Password</label><input type="password" id="loginPass" class="w-full bg-gray-800 p-3 rounded border border-gray-700" placeholder="Enter password" value="owner123"></div>
                    <button onclick="handleLogin()" class="btn-neon w-full p-3 rounded-xl"><i class="fas fa-sign-in-alt"></i> Login</button>
                    <div class="text-xs text-gray-500 text-center mt-4">
                        <p>Demo Accounts: Owner: owner/owner123 | Manager: manager/manager123</p>
                        <p>Cashier: cashier/cashier123 | Staff: staff/staff123</p>
                    </div>
                </div>
            </div>
        </div>`;
    }

    window.handleLogin = () => {
        const user = document.getElementById('loginUser')?.value;
        const pass = document.getElementById('loginPass')?.value;
        if(login(user, pass)) {
            alert(`Welcome ${appData.currentUser.name}! (${appData.currentUser.role})`);
            if (appData.systemSettings.firebaseSync && firebaseInitialized) {
                updateFirebaseStatus('connecting');
                setTimeout(() => loadAllFromFirebase(), 500);
            }
        } else {
            alert("Invalid credentials!");
        }
    };

    // ======================== OPEN SERVICES ========================
    function createOpenService(customer, vehicle, plate, items, staff, total, remarks) {
        const openService = {
            id: "OS-" + new Date().getFullYear() + "-" + (appData.nextIds.openService++),
            customer: customer || "Walk-in Customer",
            vehicle: vehicle || "N/A",
            plate: plate || "N/A",
            items: items.map(i => ({ ...i })),
            total: total,
            staff: staff || [],
            remarks: remarks || "",
            status: 'pending',
            createdAt: new Date().toISOString(),
            completedAt: null,
            paidAt: null
        };
        appData.openServices.push(openService);
        saveToLocal();
        return openService;
    }

    function updateOpenServiceStatus(id, status) {
        const service = appData.openServices.find(s => s.id === id);
        if(service) {
            service.status = status;
            if(status === 'completed') {
                service.completedAt = new Date().toISOString();
            }
            saveToLocal();
            renderApp();
        }
    }

    function completeOpenServicePayment(id, paymentMethod, cashTendered) {
        const service = appData.openServices.find(s => s.id === id);
        if(!service) return;
        
        const total = service.total;
        if(paymentMethod === "Cash" && cashTendered < total) {
            alert("Insufficient cash!");
            return;
        }
        
        const transaction = {
            id: generateReceiptNo(),
            date: new Date().toISOString(),
            customer: service.customer,
            vehicle: service.vehicle,
            plate: service.plate,
            items: service.items.map(i => ({ ...i })),
            subtotal: total,
            discount: 0,
            loyaltyUsed: 0,
            total: total,
            paymentMethod: paymentMethod,
            cashTendered: paymentMethod === "Cash" ? cashTendered : total,
            change: paymentMethod === "Cash" ? (cashTendered - total) : 0,
            remarks: service.remarks,
            isOpenService: true,
            openServiceId: service.id
        };
        
        appData.transactions.unshift(transaction);
        service.paidAt = new Date().toISOString();
        service.status = 'paid';
        
        if(service.staff && service.staff.length > 0) {
            service.staff.forEach(staff => {
                const staffMember = appData.staff.find(s => s.id === staff.id);
                if(staffMember) {
                    staffMember.totalCommission = (staffMember.totalCommission || 0) + (staff.splitCommission || 0);
                    staffMember.jobsCompleted = (staffMember.jobsCompleted || 0) + 1;
                    staffMember.salesGenerated = (staffMember.salesGenerated || 0) + total;
                }
            });
        }
        
        saveToLocal();
        alert(`✅ Payment Completed!\nTotal: ₱${total.toFixed(2)}`);
        printReceipt(transaction);
        
        appData.openServices = appData.openServices.filter(s => s.id !== id);
        renderApp();
    }

    function deleteOpenService(id) {
        if(confirm("Delete this open service?")) {
            appData.openServices = appData.openServices.filter(s => s.id !== id);
            saveToLocal();
            renderApp();
        }
    }

    function renderOpenServices() {
        if(showLogin) return renderLogin();
        
        const pendingServices = appData.openServices.filter(s => s.status === 'pending' || s.status === 'in-progress');
        const completedServices = appData.openServices.filter(s => s.status === 'completed');
        
        let pendingHtml = pendingServices.map(s => `
            <div class="open-service-card bg-gray-800 rounded-lg p-4 border border-gray-700">
                <div class="flex justify-between items-start">
                    <div>
                        <div class="font-bold text-lg">${s.customer}</div>
                        <div class="text-sm text-gray-400">${s.vehicle} | ${s.plate}</div>
                        <div class="text-xs text-gray-500">${new Date(s.createdAt).toLocaleString()}</div>
                        <div class="mt-2 space-y-1">
                            ${s.items.map(item => `<div class="text-sm"><span class="text-cyan-300">${item.serviceName}</span> (${item.variant}) x${item.quantity} - ₱${item.total}</div>`).join('')}
                        </div>
                        <div class="mt-2 text-xs text-gray-400">Staff: ${s.staff.map(st => st.name).join(', ') || 'Unassigned'}</div>
                        ${s.remarks ? `<div class="text-xs text-yellow-400 mt-1">📝 ${s.remarks}</div>` : ''}
                    </div>
                    <div class="text-right">
                        <div class="text-xl font-bold text-green-400">₱${s.total.toFixed(2)}</div>
                        <div class="mt-2 space-y-1">
                            <span class="status-badge ${s.status === 'pending' ? 'status-pending' : 'status-in-progress'}">${s.status.toUpperCase()}</span>
                            <div class="flex flex-col gap-1 mt-2">
                                <button onclick="updateOpenServiceStatus('${s.id}', 'in-progress')" class="action-btn bg-cyan-600 text-white hover:bg-cyan-700 text-xs">Start</button>
                                <button onclick="updateOpenServiceStatus('${s.id}', 'completed')" class="action-btn bg-green-600 text-white hover:bg-green-700 text-xs">Complete</button>
                                <button onclick="showPaymentModal('${s.id}')" class="action-btn bg-blue-600 text-white hover:bg-blue-700 text-xs">Pay</button>
                                <button onclick="deleteOpenService('${s.id}')" class="action-btn bg-red-600 text-white hover:bg-red-700 text-xs">Delete</button>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        `).join('');
        
        let completedHtml = completedServices.map(s => `
            <div class="open-service-card bg-gray-800/50 rounded-lg p-4 border border-green-800/30">
                <div class="flex justify-between items-start">
                    <div>
                        <div class="font-bold text-lg text-green-400">${s.customer}</div>
                        <div class="text-sm text-gray-400">${s.vehicle} | ${s.plate}</div>
                        <div class="text-xs text-gray-500">Completed: ${new Date(s.completedAt).toLocaleString()}</div>
                        <div class="mt-2 space-y-1">
                            ${s.items.map(item => `<div class="text-sm"><span class="text-cyan-300">${item.serviceName}</span> (${item.variant}) x${item.quantity}</div>`).join('')}
                        </div>
                    </div>
                    <div class="text-right">
                        <div class="text-xl font-bold text-green-400">₱${s.total.toFixed(2)}</div>
                        <span class="status-badge status-completed">COMPLETED</span>
                        <button onclick="showPaymentModal('${s.id}')" class="action-btn bg-blue-600 text-white hover:bg-blue-700 mt-2 text-xs w-full">Pay</button>
                    </div>
                </div>
            </div>
        `).join('');
        
        return `<div class="card-premium p-6">
            <div class="flex justify-between items-center mb-4">
                <h2 class="text-2xl font-bold"><i class="fas fa-clipboard-list"></i> Open Services Dashboard</h2>
                <div class="flex gap-2">
                    <span class="text-sm text-gray-400">Pending: ${pendingServices.length} | Completed: ${completedServices.length}</span>
                </div>
            </div>
            
            <div class="mb-6">
                <h3 class="text-lg font-bold text-yellow-400 mb-3"><i class="fas fa-clock"></i> Pending / In Progress</h3>
                <div class="grid lg:grid-cols-2 gap-4">
                    ${pendingHtml || '<div class="text-gray-400 text-center col-span-2 py-8">No pending services</div>'}
                </div>
            </div>
            
            <div>
                <h3 class="text-lg font-bold text-green-400 mb-3"><i class="fas fa-check-circle"></i> Completed (Awaiting Payment)</h3>
                <div class="grid lg:grid-cols-2 gap-4">
                    ${completedHtml || '<div class="text-gray-400 text-center col-span-2 py-8">No completed services awaiting payment</div>'}
                </div>
            </div>
        </div>`;
    }

    window.showPaymentModal = (id) => {
        const service = appData.openServices.find(s => s.id === id);
        if(!service) return;
        
        const paymentMethods = appData.systemSettings.defaultPaymentMethods || ['Cash', 'GCash', 'Maya', 'Credit Card', 'Bank Transfer'];
        const methodOptions = paymentMethods.map(m => `<option>${m}</option>`).join('');
        
        let modalHtml = `<div id="paymentModal" class="modal"><div class="modal-content">
            <h3 class="text-xl font-bold mb-4">Complete Payment</h3>
            <div class="space-y-3">
                <div><label class="text-gray-400">Customer</label><div class="bg-gray-800 p-2 rounded">${service.customer}</div></div>
                <div><label class="text-gray-400">Total Amount</label><div class="bg-gray-800 p-2 rounded text-2xl font-bold text-green-400">₱${service.total.toFixed(2)}</div></div>
                <div><label class="text-gray-400">Payment Method</label><select id="payMethod" class="w-full bg-gray-800 p-2 rounded">${methodOptions}</select></div>
                <div id="cashDiv"><label class="text-gray-400">Cash Tendered</label><input type="number" id="payCash" class="w-full bg-gray-800 p-2 rounded" placeholder="Enter amount" step="0.01"></div>
                <div class="flex gap-3 mt-4">
                    <button onclick="confirmPayment('${id}')" class="btn-neon flex-1 p-2 rounded">Complete Payment</button>
                    <button onclick="closePaymentModal()" class="bg-gray-700 flex-1 p-2 rounded">Cancel</button>
                </div>
            </div>
        </div></div>`;
        document.body.insertAdjacentHTML('beforeend', modalHtml);
        
        document.getElementById('payMethod').addEventListener('change', function() {
            document.getElementById('cashDiv').style.display = this.value === 'Cash' ? 'block' : 'none';
        });
    };
    
    window.confirmPayment = (id) => {
        const method = document.getElementById('payMethod')?.value || 'Cash';
        const cash = parseFloat(document.getElementById('payCash')?.value || 0);
        completeOpenServicePayment(id, method, cash);
        closePaymentModal();
    };
    
    window.closePaymentModal = () => document.getElementById('paymentModal')?.remove();

    // ======================== REPORT GENERATION ========================
    function generateDailyReport() {
        const today = new Date().toISOString().slice(0,10);
        const dailyTx = appData.transactions.filter(t => t.date.startsWith(today));
        const totalSales = dailyTx.reduce((s,t)=>s+t.total,0);
        const totalItems = dailyTx.reduce((s,t)=>s+t.items.length,0);
        const totalCustomers = [...new Set(dailyTx.map(t=>t.customer))].length;
        
        const reportData = {
            date: today,
            totalSales: totalSales,
            totalTransactions: dailyTx.length,
            totalItems: totalItems,
            totalCustomers: totalCustomers,
            transactions: dailyTx,
            paymentMethods: {}
        };
        
        dailyTx.forEach(t => {
            reportData.paymentMethods[t.paymentMethod] = (reportData.paymentMethods[t.paymentMethod] || 0) + t.total;
        });
        
        return reportData;
    }

    function generateWeeklyReport() {
        const today = new Date();
        const weekStart = new Date(today);
        weekStart.setDate(today.getDate() - today.getDay());
        const weekEnd = new Date(weekStart);
        weekEnd.setDate(weekStart.getDate() + 6);
        
        const weeklyTx = appData.transactions.filter(t => {
            const d = new Date(t.date);
            return d >= weekStart && d <= weekEnd;
        });
        
        const totalSales = weeklyTx.reduce((s,t)=>s+t.total,0);
        const reportData = {
            weekStart: weekStart.toISOString().slice(0,10),
            weekEnd: weekEnd.toISOString().slice(0,10),
            totalSales: totalSales,
            totalTransactions: weeklyTx.length,
            totalItems: weeklyTx.reduce((s,t)=>s+t.items.length,0),
            transactions: weeklyTx,
            dailyBreakdown: {}
        };
        
        weeklyTx.forEach(t => {
            const day = t.date.slice(0,10);
            if(!reportData.dailyBreakdown[day]) {
                reportData.dailyBreakdown[day] = { sales: 0, count: 0 };
            }
            reportData.dailyBreakdown[day].sales += t.total;
            reportData.dailyBreakdown[day].count += 1;
        });
        
        return reportData;
    }

    function generateMonthlyReport() {
        const now = new Date();
        const monthStart = new Date(now.getFullYear(), now.getMonth(), 1);
        const monthEnd = new Date(now.getFullYear(), now.getMonth() + 1, 0);
        
        const monthlyTx = appData.transactions.filter(t => {
            const d = new Date(t.date);
            return d >= monthStart && d <= monthEnd;
        });
        
        const totalSales = monthlyTx.reduce((s,t)=>s+t.total,0);
        const reportData = {
            month: now.toLocaleString('default', { month: 'long', year: 'numeric' }),
            monthStart: monthStart.toISOString().slice(0,10),
            monthEnd: monthEnd.toISOString().slice(0,10),
            totalSales: totalSales,
            totalTransactions: monthlyTx.length,
            totalItems: monthlyTx.reduce((s,t)=>s+t.items.length,0),
            transactions: monthlyTx,
            weeklyBreakdown: {},
            topServices: {}
        };
        
        monthlyTx.forEach(t => {
            const week = Math.ceil(new Date(t.date).getDate() / 7);
            const weekKey = `Week ${week}`;
            if(!reportData.weeklyBreakdown[weekKey]) {
                reportData.weeklyBreakdown[weekKey] = { sales: 0, count: 0 };
            }
            reportData.weeklyBreakdown[weekKey].sales += t.total;
            reportData.weeklyBreakdown[weekKey].count += 1;
            
            t.items.forEach(item => {
                reportData.topServices[item.serviceName] = (reportData.topServices[item.serviceName] || 0) + item.quantity;
            });
        });
        
        return reportData;
    }

    function generateCommissionReport() {
        const commissionData = appData.staff.map(s => {
            const totalCommission = s.totalCommission || 0;
            const dailyCommission = appData.transactions.filter(t => {
                return t.staffCommissions && t.staffCommissions[s.id];
            }).reduce((sum, t) => {
                return sum + (t.staffCommissions[s.id]?.commission || 0);
            }, 0);
            
            return {
                name: s.name,
                position: s.position,
                commissionTier: s.commissionTier,
                jobsCompleted: s.jobsCompleted || 0,
                salesGenerated: s.salesGenerated || 0,
                dailyCommission: dailyCommission,
                totalCommission: totalCommission
            };
        });
        
        return commissionData;
    }

    function generateEnhancedCommissionReport() {
        const reportData = [];
        const today = new Date().toISOString().slice(0,10);
        
        appData.staff.forEach(staff => {
            const staffTransactions = appData.transactions.filter(t => {
                return t.staffCommissions && t.staffCommissions[staff.id];
            });
            
            const dailyBreakdown = {};
            staffTransactions.forEach(t => {
                const date = t.date.slice(0,10);
                if (!dailyBreakdown[date]) {
                    dailyBreakdown[date] = { total: 0, transactions: [] };
                }
                const commission = t.staffCommissions[staff.id]?.commission || 0;
                dailyBreakdown[date].total += commission;
                dailyBreakdown[date].transactions.push({
                    id: t.id,
                    customer: t.customer,
                    amount: commission,
                    totalSale: t.total
                });
            });
            
            const totalCommission = staffTransactions.reduce((sum, t) => {
                return sum + (t.staffCommissions[staff.id]?.commission || 0);
            }, 0);
            
            const dailyCommission = staffTransactions.filter(t => t.date.startsWith(today))
                .reduce((sum, t) => sum + (t.staffCommissions[staff.id]?.commission || 0), 0);
            
            reportData.push({
                staffId: staff.id,
                name: staff.name,
                position: staff.position,
                commissionTier: staff.commissionTier,
                jobsCompleted: staff.jobsCompleted || 0,
                salesGenerated: staff.salesGenerated || 0,
                dailyCommission: dailyCommission,
                totalCommission: totalCommission,
                dailyBreakdown: dailyBreakdown,
                transactionCount: staffTransactions.length
            });
        });
        
        return reportData;
    }

    // ======================== EXCEL EXPORT FUNCTIONS ========================
    function exportToExcel(reportType) {
        const wb = XLSX.utils.book_new();
        let data = [];
        let sheetName = '';
        
        if (reportType === 'daily') {
            const report = generateDailyReport();
            const totalCommission = appData.staff.reduce((s, st) => s + (st.totalCommission || 0), 0);
            
            data = [
                ['8068 CAR SQUARE - DAILY SALES REPORT'],
                ['Date:', new Date(report.date).toLocaleDateString()],
                [''],
                ['SUMMARY'],
                ['Total Sales', `₱${report.totalSales.toFixed(2)}`],
                ['Total Transactions', report.totalTransactions],
                ['Total Items Sold', report.totalItems],
                ['Total Customers', report.totalCustomers],
                ['Total Commission Paid', `₱${totalCommission.toFixed(2)}`],
                [''],
                ['PAYMENT METHOD BREAKDOWN'],
                ['Payment Method', 'Amount'],
                ...Object.entries(report.paymentMethods).map(([method, amount]) => [method, `₱${amount.toFixed(2)}`]),
                [''],
                ['TRANSACTION DETAILS'],
                ['Receipt', 'Date & Time', 'Customer', 'Vehicle', 'Plate', 'Service', 'Size', 'Price', 'Staff', 'Payment', 'Total'],
                ...report.transactions.flatMap(t => 
                    t.items.map(item => [
                        t.id,
                        new Date(t.date).toLocaleString(),
                        t.customer,
                        t.vehicle,
                        t.plate,
                        item.serviceName,
                        item.variant,
                        `₱${item.price.toFixed(2)}`,
                        item.assignedStaff?.map(s => s.name).join(', ') || 'Unassigned',
                        t.paymentMethod,
                        `₱${t.total.toFixed(2)}`
                    ])
                )
            ];
            sheetName = 'Daily Report';
        } else if (reportType === 'weekly') {
            const report = generateWeeklyReport();
            const totalCommission = appData.staff.reduce((s, st) => s + (st.totalCommission || 0), 0);
            
            data = [
                ['8068 CAR SQUARE - WEEKLY SALES REPORT'],
                ['Week:', `${new Date(report.weekStart).toLocaleDateString()} - ${new Date(report.weekEnd).toLocaleDateString()}`],
                [''],
                ['SUMMARY'],
                ['Total Sales', `₱${report.totalSales.toFixed(2)}`],
                ['Total Transactions', report.totalTransactions],
                ['Total Items Sold', report.totalItems],
                ['Total Commission Paid', `₱${totalCommission.toFixed(2)}`],
                [''],
                ['DAILY BREAKDOWN'],
                ['Date', 'Sales', 'Transactions'],
                ...Object.entries(report.dailyBreakdown).map(([date, data]) => [
                    new Date(date).toLocaleDateString(),
                    `₱${data.sales.toFixed(2)}`,
                    data.count
                ]),
                [''],
                ['TRANSACTION DETAILS'],
                ['Receipt', 'Date & Time', 'Customer', 'Vehicle', 'Plate', 'Service', 'Size', 'Price', 'Staff', 'Payment', 'Total'],
                ...report.transactions.flatMap(t => 
                    t.items.map(item => [
                        t.id,
                        new Date(t.date).toLocaleString(),
                        t.customer,
                        t.vehicle,
                        t.plate,
                        item.serviceName,
                        item.variant,
                        `₱${item.price.toFixed(2)}`,
                        item.assignedStaff?.map(s => s.name).join(', ') || 'Unassigned',
                        t.paymentMethod,
                        `₱${t.total.toFixed(2)}`
                    ])
                )
            ];
            sheetName = 'Weekly Report';
        } else if (reportType === 'monthly') {
            const report = generateMonthlyReport();
            const totalCommission = appData.staff.reduce((s, st) => s + (st.totalCommission || 0), 0);
            
            data = [
                ['8068 CAR SQUARE - MONTHLY SALES REPORT'],
                ['Month:', report.month],
                [''],
                ['SUMMARY'],
                ['Total Sales', `₱${report.totalSales.toFixed(2)}`],
                ['Total Transactions', report.totalTransactions],
                ['Total Items Sold', report.totalItems],
                ['Total Commission Paid', `₱${totalCommission.toFixed(2)}`],
                [''],
                ['WEEKLY BREAKDOWN'],
                ['Week', 'Sales', 'Transactions'],
                ...Object.entries(report.weeklyBreakdown).map(([week, data]) => [
                    week,
                    `₱${data.sales.toFixed(2)}`,
                    data.count
                ]),
                [''],
                ['TOP SERVICES'],
                ['Service', 'Quantity Sold'],
                ...Object.entries(report.topServices).sort((a, b) => b[1] - a[1]),
                [''],
                ['TRANSACTION DETAILS'],
                ['Receipt', 'Date & Time', 'Customer', 'Vehicle', 'Plate', 'Service', 'Size', 'Price', 'Staff', 'Payment', 'Total'],
                ...report.transactions.flatMap(t => 
                    t.items.map(item => [
                        t.id,
                        new Date(t.date).toLocaleString(),
                        t.customer,
                        t.vehicle,
                        t.plate,
                        item.serviceName,
                        item.variant,
                        `₱${item.price.toFixed(2)}`,
                        item.assignedStaff?.map(s => s.name).join(', ') || 'Unassigned',
                        t.paymentMethod,
                        `₱${t.total.toFixed(2)}`
                    ])
                )
            ];
            sheetName = 'Monthly Report';
        } else if (reportType === 'commission') {
            const report = generateCommissionReport();
            const totalCommission = report.reduce((sum, s) => sum + s.totalCommission, 0);
            
            data = [
                ['8068 CAR SQUARE - EMPLOYEE COMMISSION REPORT'],
                ['Date Generated:', new Date().toLocaleString()],
                [''],
                ['COMMISSION SUMMARY'],
                ['Staff Name', 'Position', 'Commission Tier', 'Jobs Completed', 'Sales Generated', 'Daily Commission', 'Total Commission'],
                ...report.map(s => [
                    s.name,
                    s.position,
                    `${s.commissionTier}%`,
                    s.jobsCompleted,
                    `₱${s.salesGenerated.toFixed(2)}`,
                    `₱${s.dailyCommission.toFixed(2)}`,
                    `₱${s.totalCommission.toFixed(2)}`
                ]),
                [''],
                ['GRAND TOTAL', '', '', '', '', `₱${report.reduce((sum, s) => sum + s.dailyCommission, 0).toFixed(2)}`, `₱${totalCommission.toFixed(2)}`]
            ];
            sheetName = 'Commission Report';
        }
        
        const ws = XLSX.utils.aoa_to_sheet(data);
        
        const colWidths = [];
        if (data.length > 0) {
            const firstRow = data[0];
            for (let i = 0; i < firstRow.length; i++) {
                let maxWidth = 12;
                for (let row of data) {
                    if (row[i] !== undefined && row[i] !== null) {
                        const val = String(row[i]);
                        maxWidth = Math.max(maxWidth, val.length + 3);
                    }
                }
                colWidths.push({ wch: Math.min(maxWidth, 45) });
            }
        }
        ws['!cols'] = colWidths;
        
        const range = XLSX.utils.decode_range(ws['!ref']);
        for (let R = range.s.r; R <= range.e.r; R++) {
            for (let C = range.s.c; C <= range.e.c; C++) {
                const addr = XLSX.utils.encode_cell({ r: R, c: C });
                if (!ws[addr]) continue;
                if (!ws[addr].s) ws[addr].s = {};
                ws[addr].s.alignment = { wrapText: true, vertical: 'top' };
            }
        }
        
        XLSX.utils.book_append_sheet(wb, ws, sheetName);
        
        const fileName = `${reportType}_report_${new Date().toISOString().slice(0,10)}.xlsx`;
        XLSX.writeFile(wb, fileName);
        alert(`✅ Excel report exported as: ${fileName}`);
    }

    function exportCommissionToExcel() {
        const report = generateEnhancedCommissionReport();
        const wb = XLSX.utils.book_new();
        
        // Main Summary Sheet
        let summaryData = [
            ['8068 CAR SQUARE - EMPLOYEE COMMISSION REPORT'],
            ['Date Generated:', new Date().toLocaleString()],
            [''],
            ['COMMISSION SUMMARY'],
            ['Staff Name', 'Position', 'Commission Tier', 'Jobs Completed', 'Sales Generated', 'Daily Commission', 'Total Commission'],
            ...report.map(s => [
                s.name,
                s.position,
                `${s.commissionTier}%`,
                s.jobsCompleted,
                `₱${s.salesGenerated.toFixed(2)}`,
                `₱${s.dailyCommission.toFixed(2)}`,
                `₱${s.totalCommission.toFixed(2)}`
            ]),
            [''],
            ['GRAND TOTAL', '', '', '', '', `₱${report.reduce((sum, s) => sum + s.dailyCommission, 0).toFixed(2)}`, `₱${report.reduce((sum, s) => sum + s.totalCommission, 0).toFixed(2)}`]
        ];
        
        const ws1 = XLSX.utils.aoa_to_sheet(summaryData);
        const colWidths1 = [];
        for (let i = 0; i < summaryData[0].length; i++) {
            let maxWidth = 12;
            for (let row of summaryData) {
                if (row[i] !== undefined && row[i] !== null) {
                    const val = String(row[i]);
                    maxWidth = Math.max(maxWidth, val.length + 3);
                }
            }
            colWidths1.push({ wch: Math.min(maxWidth, 45) });
        }
        ws1['!cols'] = colWidths1;
        XLSX.utils.book_append_sheet(wb, ws1, 'Summary');
        
        // Daily Breakdown Sheet
        let dailyData = [
            ['DAILY COMMISSION BREAKDOWN'],
            [''],
            ['Staff', 'Date', 'Transactions', 'Commission', 'Customers', 'Receipts']
        ];
        
        report.forEach(s => {
            const dates = Object.keys(s.dailyBreakdown).sort();
            dates.forEach(date => {
                const dayData = s.dailyBreakdown[date];
                const customers = [...new Set(dayData.transactions.map(t => t.customer))];
                const receipts = dayData.transactions.map(t => t.id).join(', ');
                dailyData.push([
                    s.name,
                    new Date(date).toLocaleDateString(),
                    dayData.transactions.length,
                    `₱${dayData.total.toFixed(2)}`,
                    customers.length,
                    receipts
                ]);
            });
        });
        
        const ws2 = XLSX.utils.aoa_to_sheet(dailyData);
        const colWidths2 = [];
        for (let i = 0; i < dailyData[0].length; i++) {
            let maxWidth = 12;
            for (let row of dailyData) {
                if (row[i] !== undefined && row[i] !== null) {
                    const val = String(row[i]);
                    maxWidth = Math.max(maxWidth, val.length + 3);
                }
            }
            colWidths2.push({ wch: Math.min(maxWidth, 45) });
        }
        ws2['!cols'] = colWidths2;
        
        const range2 = XLSX.utils.decode_range(ws2['!ref']);
        for (let R = range2.s.r; R <= range2.e.r; R++) {
            for (let C = range2.s.c; C <= range2.e.c; C++) {
                const addr = XLSX.utils.encode_cell({ r: R, c: C });
                if (!ws2[addr]) continue;
                if (!ws2[addr].s) ws2[addr].s = {};
                ws2[addr].s.alignment = { wrapText: true, vertical: 'top' };
            }
        }
        XLSX.utils.book_append_sheet(wb, ws2, 'Daily Breakdown');
        
        // Detailed Transactions Sheet
        let txData = [
            ['TRANSACTION DETAILS'],
            [''],
            ['Staff', 'Date', 'Receipt', 'Customer', 'Commission', 'Total Sale']
        ];
        
        appData.transactions.forEach(t => {
            if (t.staffCommissions) {
                Object.entries(t.staffCommissions).forEach(([staffId, data]) => {
                    const staff = appData.staff.find(s => s.id === parseInt(staffId));
                    if (staff) {
                        txData.push([
                            staff.name,
                            new Date(t.date).toLocaleDateString(),
                            t.id,
                            t.customer,
                            `₱${data.commission.toFixed(2)}`,
                            `₱${t.total.toFixed(2)}`
                        ]);
                    }
                });
            }
        });
        
        const ws3 = XLSX.utils.aoa_to_sheet(txData);
        const colWidths3 = [];
        for (let i = 0; i < txData[0].length; i++) {
            let maxWidth = 12;
            for (let row of txData) {
                if (row[i] !== undefined && row[i] !== null) {
                    const val = String(row[i]);
                    maxWidth = Math.max(maxWidth, val.length + 3);
                }
            }
            colWidths3.push({ wch: Math.min(maxWidth, 45) });
        }
        ws3['!cols'] = colWidths3;
        
        const range3 = XLSX.utils.decode_range(ws3['!ref']);
        for (let R = range3.s.r; R <= range3.e.r; R++) {
            for (let C = range3.s.c; C <= range3.e.c; C++) {
                const addr = XLSX.utils.encode_cell({ r: R, c: C });
                if (!ws3[addr]) continue;
                if (!ws3[addr].s) ws3[addr].s = {};
                ws3[addr].s.alignment = { wrapText: true, vertical: 'top' };
            }
        }
        XLSX.utils.book_append_sheet(wb, ws3, 'Transactions');
        
        const fileName = `commission_report_${new Date().toISOString().slice(0,10)}.xlsx`;
        XLSX.writeFile(wb, fileName);
        alert(`✅ Enhanced commission report exported as: ${fileName}`);
    }

    // ======================== STAFF FUNCTIONS ========================
    function renderStaffDashboard() {
        if(showLogin) return renderLogin();
        const totalCommission = appData.staff.reduce((s, st) => s + (st.totalCommission || 0), 0);
        const totalJobs = appData.staff.reduce((s, st) => s + (st.jobsCompleted || 0), 0);
        const totalSales = appData.staff.reduce((s, st) => s + (st.salesGenerated || 0), 0);
        
        return `<div class="card-premium p-6">
            <div class="flex justify-between items-center mb-4">
                <h2 class="text-2xl font-bold"><i class="fas fa-id-badge"></i> Staff Dashboard</h2>
                <div class="flex gap-2 flex-wrap">
                    <button onclick="exportStaffCSV()" class="excel-export-btn px-4 py-2 rounded text-sm text-white"><i class="fas fa-file-excel"></i> Export CSV</button>
                    <button onclick="printStaffDashboard()" class="bg-cyan-600 px-4 py-2 rounded text-sm"><i class="fas fa-print"></i> Print Dashboard</button>
                    <button onclick="addStaffPopup()" class="bg-green-600 px-4 py-2 rounded text-sm"><i class="fas fa-plus"></i> Add Staff</button>
                </div>
            </div>
            
            <div class="grid grid-cols-2 md:grid-cols-4 gap-4 mb-6">
                <div class="bg-gray-800 p-4 rounded-lg text-center">
                    <div class="text-gray-400 text-sm">Total Staff</div>
                    <div class="text-2xl font-bold text-cyan-400">${appData.staff.length}</div>
                </div>
                <div class="bg-gray-800 p-4 rounded-lg text-center">
                    <div class="text-gray-400 text-sm">Total Jobs Completed</div>
                    <div class="text-2xl font-bold text-green-400">${totalJobs}</div>
                </div>
                <div class="bg-gray-800 p-4 rounded-lg text-center">
                    <div class="text-gray-400 text-sm">Total Sales Generated</div>
                    <div class="text-2xl font-bold text-yellow-400">₱${totalSales.toFixed(2)}</div>
                </div>
                <div class="bg-gray-800 p-4 rounded-lg text-center">
                    <div class="text-gray-400 text-sm">Total Commissions Paid</div>
                    <div class="text-2xl font-bold text-purple-400">₱${totalCommission.toFixed(2)}</div>
                </div>
            </div>
            
            <div class="overflow-x-auto">
                <table class="w-full text-sm" id="staffTable">
                    <thead class="border-b-2 border-cyan-400">
                        <tr>
                            <th class="text-left py-2">#</th>
                            <th class="text-left py-2">Staff Name</th>
                            <th class="text-left py-2">Position</th>
                            <th class="text-center py-2">Tier</th>
                            <th class="text-center py-2">Jobs Done</th>
                            <th class="text-right py-2">Sales Generated</th>
                            <th class="text-right py-2">Commission</th>
                            <th class="text-center py-2">Status</th>
                            <th class="text-center py-2">Actions</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${appData.staff.map((s, idx) => `
                            <tr class="border-b border-gray-700 hover:bg-gray-800 transition">
                                <td class="py-2">${idx + 1}</td>
                                <td class="editable-field cursor-pointer" onclick="editStaffName(${s.id})">${s.name}</td>
                                <td class="editable-field cursor-pointer" onclick="editStaffPosition(${s.id})">${s.position}</td>
                                <td class="text-center"><span class="commission-badge editable-field cursor-pointer" onclick="editStaffCommission(${s.id})">${s.commissionTier}%</span></td>
                                <td class="text-center editable-field cursor-pointer" onclick="editStaffJobs(${s.id})">${s.jobsCompleted || 0}</td>
                                <td class="text-right editable-field cursor-pointer" onclick="editStaffSales(${s.id})">₱${(s.salesGenerated || 0).toFixed(2)}</td>
                                <td class="text-right text-green-400">₱${(s.totalCommission || 0).toFixed(2)}</td>
                                <td class="text-center"><button onclick="toggleStaffStatus(${s.id})" class="text-${s.isActive ? 'green' : 'red'}-400 text-sm">${s.isActive ? 'Active' : 'Inactive'}</button></td>
                                <td class="text-center">
                                    <button onclick="deleteStaffMember(${s.id})" class="text-red-400"><i class="fas fa-trash"></i></button>
                                </td>
                            </tr>
                        `).join('')}
                        ${appData.staff.length === 0 ? '<tr><td colspan="9" class="text-center py-4 text-gray-400">No staff members added yet</td></tr>' : ''}
                    </tbody>
                    <tfoot class="border-t-2 border-cyan-400 bg-gray-800 font-bold">
                        <tr>
                            <td colspan="4" class="py-2 text-right">TOTALS:</td>
                            <td class="text-center">${totalJobs}</td>
                            <td class="text-right">₱${totalSales.toFixed(2)}</td>
                            <td class="text-right text-green-400">₱${totalCommission.toFixed(2)}</td>
                            <td colspan="2"></td>
                        </tr>
                    </tfoot>
                </table>
            </div>
        </div>`;
    }

    window.editStaffName = (id) => {
        const staff = appData.staff.find(s => s.id === id);
        if(staff) {
            const newName = prompt("Edit staff name:", staff.name);
            if(newName) { staff.name = newName; saveToLocal(); renderApp(); }
        }
    };
    window.editStaffPosition = (id) => {
        const staff = appData.staff.find(s => s.id === id);
        if(staff) {
            const newPos = prompt("Edit position:", staff.position);
            if(newPos) { staff.position = newPos; saveToLocal(); renderApp(); }
        }
    };
    window.editStaffCommission = (id) => {
        const staff = appData.staff.find(s => s.id === id);
        if(staff) {
            const newTier = parseInt(prompt("Edit commission tier (30, 35, 40):", staff.commissionTier));
            if([30,35,40].includes(newTier)) { staff.commissionTier = newTier; saveToLocal(); renderApp(); }
        }
    };
    window.editStaffJobs = (id) => {
        const staff = appData.staff.find(s => s.id === id);
        if(staff) {
            const newJobs = parseInt(prompt("Edit jobs completed:", staff.jobsCompleted || 0));
            if(!isNaN(newJobs) && newJobs >= 0) { staff.jobsCompleted = newJobs; saveToLocal(); renderApp(); }
        }
    };
    window.editStaffSales = (id) => {
        const staff = appData.staff.find(s => s.id === id);
        if(staff) {
            const newSales = parseFloat(prompt("Edit sales generated:", staff.salesGenerated || 0));
            if(!isNaN(newSales) && newSales >= 0) { staff.salesGenerated = newSales; saveToLocal(); renderApp(); }
        }
    };
    window.toggleStaffStatus = (id) => {
        const staff = appData.staff.find(s => s.id === id);
        if(staff) { staff.isActive = !staff.isActive; saveToLocal(); renderApp(); }
    };
    window.deleteStaffMember = (id) => {
        if(confirm("Delete this staff member?")) {
            appData.staff = appData.staff.filter(s => s.id !== id);
            saveToLocal(); renderApp();
        }
    };
    window.addStaffPopup = () => {
        let name = prompt("Staff Name:"); if(!name) return;
        let pos = prompt("Position:") || "Staff";
        let tier = parseInt(prompt("Commission Tier (30, 35, 40):", "35"));
        if([30,35,40].includes(tier)) {
            appData.staff.push({ id: Date.now(), name, position: pos, contact: "", salesGenerated: 0, jobsCompleted: 0, commissionTier: tier, isActive: true, totalCommission: 0 });
            saveToLocal(); renderApp();
        }
    };

    window.exportStaffCSV = () => {
        let csv = "Staff Name,Position,Commission Tier,Jobs Completed,Sales Generated,Total Commission,Status\n";
        appData.staff.forEach(s => {
            csv += `"${s.name}","${s.position}","${s.commissionTier}%",${s.jobsCompleted || 0},${(s.salesGenerated || 0).toFixed(2)},${(s.totalCommission || 0).toFixed(2)},"${s.isActive ? 'Active' : 'Inactive'}"\n`;
        });
        const totalCommission = appData.staff.reduce((sum, s) => sum + (s.totalCommission || 0), 0);
        const totalSales = appData.staff.reduce((sum, s) => sum + (s.salesGenerated || 0), 0);
        const totalJobs = appData.staff.reduce((sum, s) => sum + (s.jobsCompleted || 0), 0);
        csv += `\n"TOTALS","","",${totalJobs},${totalSales.toFixed(2)},${totalCommission.toFixed(2)},""\n`;
        
        const blob = new Blob([csv], {type: "text/csv"});
        const a = document.createElement("a");
        a.href = URL.createObjectURL(blob);
        a.download = `staff_report_${new Date().toISOString().slice(0,10)}.csv`;
        a.click();
        alert("✅ Staff report exported as CSV!");
    };

    window.printStaffDashboard = () => {
        const printContent = document.querySelector('.card-premium').cloneNode(true);
        const win = window.open('', '_blank');
        const logoHtml = appData.systemSettings.companyLogo ? 
            `<div class="center"><img src="${appData.systemSettings.companyLogo}" style="max-width:100px;max-height:70px;object-fit:contain;"></div>` : '';
        win.document.write(`<html><head><title>Staff Dashboard</title><style>
            body{font-family:monospace;padding:20px;background:white;color:black;}
            .center{text-align:center;}
            table{width:100%;border-collapse:collapse;}
            th,td{border:1px solid #ddd;padding:8px;text-align:left;}
            th{background:#f0f0f0;}
            .header{margin-bottom:20px;}
        </style></head><body>
            ${logoHtml}
            <div class="header"><h2>${appData.systemSettings.companyName}</h2><p>Staff Dashboard - ${new Date().toLocaleDateString()}</p></div>
            ${printContent.innerHTML}
        </body></html>`);
        win.document.close();
        win.print();
    };

    // ======================== REPORTS DASHBOARD ========================
    function renderReports() {
        if(showLogin) return renderLogin();
        const dailyReport = generateDailyReport();
        const weeklyReport = generateWeeklyReport();
        const monthlyReport = generateMonthlyReport();
        const enhancedCommission = generateEnhancedCommissionReport();
        const totalDailyCommission = enhancedCommission.reduce((sum, s) => sum + s.dailyCommission, 0);
        const totalCommission = enhancedCommission.reduce((sum, s) => sum + s.totalCommission, 0);
        
        return `<div class="card-premium p-6">
            <h2 class="text-2xl font-bold mb-4"><i class="fas fa-file-alt"></i> Reports Dashboard</h2>
            
            <div class="grid lg:grid-cols-4 gap-4 mb-6">
                <div class="bg-gray-800 p-4 rounded-lg text-center">
                    <div class="text-gray-400 text-sm">Total Transactions</div>
                    <div class="text-2xl font-bold text-cyan-400">${appData.transactions.length}</div>
                </div>
                <div class="bg-gray-800 p-4 rounded-lg text-center">
                    <div class="text-gray-400 text-sm">Total Revenue</div>
                    <div class="text-2xl font-bold text-green-400">₱${appData.transactions.reduce((a,b)=>a+b.total,0).toFixed(2)}</div>
                </div>
                <div class="bg-gray-800 p-4 rounded-lg text-center">
                    <div class="text-gray-400 text-sm">Total Customers</div>
                    <div class="text-2xl font-bold text-yellow-400">${appData.customers.length}</div>
                </div>
                <div class="bg-gray-800 p-4 rounded-lg text-center">
                    <div class="text-gray-400 text-sm">Total Staff</div>
                    <div class="text-2xl font-bold text-purple-400">${appData.staff.length}</div>
                </div>
            </div>
            
            <div class="grid lg:grid-cols-3 gap-4 mb-6">
                <div class="bg-gray-800 p-4 rounded-lg">
                    <h3 class="text-lg font-bold text-cyan-400 mb-2"><i class="fas fa-calendar-day"></i> Daily Summary</h3>
                    <div class="text-sm space-y-1">
                        <div class="flex justify-between"><span>Date:</span><span>${dailyReport.date}</span></div>
                        <div class="flex justify-between"><span>Sales:</span><span class="text-green-400">₱${dailyReport.totalSales.toFixed(2)}</span></div>
                        <div class="flex justify-between"><span>Transactions:</span><span>${dailyReport.totalTransactions}</span></div>
                        <div class="flex justify-between"><span>Customers:</span><span>${dailyReport.totalCustomers}</span></div>
                    </div>
                    <button onclick="exportToExcel('daily')" class="excel-export-btn w-full mt-2 p-2 rounded text-sm text-white"><i class="fas fa-file-excel"></i> Export Daily Excel</button>
                </div>
                <div class="bg-gray-800 p-4 rounded-lg">
                    <h3 class="text-lg font-bold text-green-400 mb-2"><i class="fas fa-calendar-week"></i> Weekly Summary</h3>
                    <div class="text-sm space-y-1">
                        <div class="flex justify-between"><span>Week:</span><span>${weeklyReport.weekStart} - ${weeklyReport.weekEnd}</span></div>
                        <div class="flex justify-between"><span>Sales:</span><span class="text-green-400">₱${weeklyReport.totalSales.toFixed(2)}</span></div>
                        <div class="flex justify-between"><span>Transactions:</span><span>${weeklyReport.totalTransactions}</span></div>
                        <div class="flex justify-between"><span>Items Sold:</span><span>${weeklyReport.totalItems}</span></div>
                    </div>
                    <button onclick="exportToExcel('weekly')" class="excel-export-btn w-full mt-2 p-2 rounded text-sm text-white"><i class="fas fa-file-excel"></i> Export Weekly Excel</button>
                </div>
                <div class="bg-gray-800 p-4 rounded-lg">
                    <h3 class="text-lg font-bold text-yellow-400 mb-2"><i class="fas fa-calendar-alt"></i> Monthly Summary</h3>
                    <div class="text-sm space-y-1">
                        <div class="flex justify-between"><span>Month:</span><span>${monthlyReport.month}</span></div>
                        <div class="flex justify-between"><span>Sales:</span><span class="text-green-400">₱${monthlyReport.totalSales.toFixed(2)}</span></div>
                        <div class="flex justify-between"><span>Transactions:</span><span>${monthlyReport.totalTransactions}</span></div>
                        <div class="flex justify-between"><span>Items Sold:</span><span>${monthlyReport.totalItems}</span></div>
                    </div>
                    <button onclick="exportToExcel('monthly')" class="excel-export-btn w-full mt-2 p-2 rounded text-sm text-white"><i class="fas fa-file-excel"></i> Export Monthly Excel</button>
                </div>
            </div>
            
            <div class="bg-gray-800 p-4 rounded-lg">
                <div class="flex justify-between items-center mb-3">
                    <h3 class="text-lg font-bold text-purple-400"><i class="fas fa-file-invoice-dollar"></i> Enhanced Commission Report</h3>
                    <div class="flex gap-2">
                        <button onclick="exportCommissionToExcel()" class="excel-export-btn px-4 py-2 rounded text-sm text-white"><i class="fas fa-file-excel"></i> Export Enhanced Excel</button>
                        <button onclick="exportToExcel('commission')" class="bg-blue-600 px-4 py-2 rounded text-sm text-white"><i class="fas fa-file-excel"></i> Basic Report</button>
                    </div>
                </div>
                <div class="grid grid-cols-2 md:grid-cols-4 gap-2 mb-3">
                    <div class="bg-gray-700 p-2 rounded text-center text-xs">
                        <div class="text-gray-400">Daily Commission</div>
                        <div class="text-lg font-bold text-cyan-400">₱${totalDailyCommission.toFixed(2)}</div>
                    </div>
                    <div class="bg-gray-700 p-2 rounded text-center text-xs">
                        <div class="text-gray-400">Total Commission</div>
                        <div class="text-lg font-bold text-green-400">₱${totalCommission.toFixed(2)}</div>
                    </div>
                    <div class="bg-gray-700 p-2 rounded text-center text-xs">
                        <div class="text-gray-400">Total Staff</div>
                        <div class="text-lg font-bold text-yellow-400">${appData.staff.length}</div>
                    </div>
                    <div class="bg-gray-700 p-2 rounded text-center text-xs">
                        <div class="text-gray-400">Total Jobs</div>
                        <div class="text-lg font-bold text-purple-400">${appData.staff.reduce((sum, s) => sum + (s.jobsCompleted || 0), 0)}</div>
                    </div>
                </div>
                <div class="overflow-x-auto max-h-80 overflow-y-auto">
                    <table class="w-full text-sm">
                        <thead class="sticky top-0 bg-gray-900 border-b border-purple-400">
                            <tr><th>Staff Name</th><th>Position</th><th>Tier</th><th>Jobs</th><th>Sales</th><th>Daily Commission</th><th>Total Commission</th><th>Daily Breakdown</th></tr>
                        </thead>
                        <tbody>
                            ${enhancedCommission.map(s => {
                                const dates = Object.keys(s.dailyBreakdown).sort().slice(0,3);
                                const breakdownHtml = dates.map(date => 
                                    `<div class="text-xs text-gray-400">${new Date(date).toLocaleDateString()}: ₱${s.dailyBreakdown[date].total.toFixed(2)} (${s.dailyBreakdown[date].transactions.length} tx)</div>`
                                ).join('');
                                return `<tr class="border-b border-gray-700 commission-detail-row">
                                    <td class="py-2 font-semibold">${s.name}</td>
                                    <td>${s.position}</td>
                                    <td><span class="commission-badge">${s.commissionTier}%</span></td>
                                    <td>${s.jobsCompleted}</td>
                                    <td>₱${s.salesGenerated.toFixed(2)}</td>
                                    <td class="text-cyan-400 font-bold">₱${s.dailyCommission.toFixed(2)}</td>
                                    <td class="text-green-400 font-bold">₱${s.totalCommission.toFixed(2)}</td>
                                    <td class="text-xs">${breakdownHtml || 'No daily data'}</td>
                                </tr>`;
                            }).join('')}
                            <tr class="border-t-2 border-purple-400 font-bold">
                                <td colspan="5" class="py-2 text-right">GRAND TOTAL:</td>
                                <td class="text-cyan-400">₱${totalDailyCommission.toFixed(2)}</td>
                                <td class="text-green-400">₱${totalCommission.toFixed(2)}</td>
                                <td></td>
                            </tr>
                        </tbody>
                    </table>
                </div>
            </div>
        </div>`;
    }

    // ======================== DASHBOARD ========================
    function renderDashboard() {
        if(showLogin) return renderLogin();
        const todaySales = appData.transactions.filter(t=>new Date(t.date).toDateString() === new Date().toDateString()).reduce((s,t)=>s+t.total,0);
        const weeklySales = appData.transactions.filter(t=>{ const d=new Date(t.date); const today=new Date(); const weekStart=new Date(today); weekStart.setDate(today.getDate()-today.getDay()); return d>=weekStart; }).reduce((s,t)=>s+t.total,0);
        const monthlySales = appData.transactions.filter(t=>new Date(t.date).getMonth() === new Date().getMonth()).reduce((s,t)=>s+t.total,0);
        const openCount = appData.openServices.filter(s => s.status === 'pending' || s.status === 'in-progress').length;
        const lowStock = appData.inventory.filter(i => i.quantity <= i.reorderLevel).length;

        const today = new Date().toISOString().slice(0,10);
        const staffDailyCommissions = appData.staff.map(s => {
            const dailyCommission = appData.transactions.filter(t => {
                return t.date.startsWith(today) && t.staffCommissions && t.staffCommissions[s.id];
            }).reduce((sum, t) => {
                return sum + (t.staffCommissions[s.id]?.commission || 0);
            }, 0);
            return { ...s, dailyCommission: dailyCommission };
        }).filter(s => s.dailyCommission > 0 || s.isActive);

        const totalDailyCommission = staffDailyCommissions.reduce((sum, s) => sum + s.dailyCommission, 0);

        const recentCommissions = appData.transactions
            .filter(t => t.date.startsWith(today) && t.staffCommissions)
            .slice(0, 5)
            .map(t => {
                const staffNames = Object.values(t.staffCommissions).map(s => s.name).join(', ');
                const totalComm = Object.values(t.staffCommissions).reduce((sum, s) => sum + s.commission, 0);
                return { time: t.date, staff: staffNames, amount: totalComm, customer: t.customer };
            });

        const serviceCount = {};
        appData.transactions.forEach(t => { t.items.forEach(item => { serviceCount[item.serviceName] = (serviceCount[item.serviceName] || 0) + item.quantity; }); });
        const sortedServices = Object.entries(serviceCount).sort((a,b) => b[1] - a[1]).slice(0,8);
        const maxCount = sortedServices.length > 0 ? sortedServices[0][1] : 1;
        
        let barHtml = sortedServices.map(([name, count]) => {
            const pct = (count / maxCount * 100);
            const colors = ['#00d4ff', '#00ff88', '#ffaa00', '#ff4444', '#aa66ff', '#ff66aa', '#66ffcc', '#ffcc44'];
            const color = colors[sortedServices.indexOf([name, count]) % colors.length];
            return `<div class="bar-item"><span class="bar-label text-sm">${name}</span><div class="bar-track"><div class="bar-fill" style="width:${pct}%;background:${color}"></div></div><span class="bar-value text-sm">${count}</span></div>`;
        }).join('');

        setTimeout(() => {
            const ctx = document.getElementById('salesChart')?.getContext('2d');
            if(ctx) { new Chart(ctx, { type: 'line', data: { labels: ['Mon','Tue','Wed','Thu','Fri','Sat','Sun'], datasets: [{ label: 'Sales (₱)', data: [1500,2300,3200,2800,4100,3900,4800], borderColor: '#00d4ff', tension: 0.3, fill: true, backgroundColor: 'rgba(0,212,255,0.1)' }] } }); }
        }, 100);
        
        const logoHtml = appData.systemSettings.companyLogo ? 
            `<img src="${appData.systemSettings.companyLogo}" class="company-logo inline-block mr-2" alt="Logo">` : '';
        
        const firebaseStatus = firebaseInitialized ? 'connected' : 'disconnected';
        
        let commissionHtml = staffDailyCommissions.map(s => `
            <div class="commission-item flex justify-between items-center py-2 border-b border-gray-700/50">
                <div class="flex items-center gap-2">
                    <span class="text-sm font-medium">${s.name}</span>
                    <span class="commission-badge text-xs">${s.commissionTier}%</span>
                </div>
                <div class="flex items-center gap-3">
                    <span class="text-xs text-gray-400">${s.jobsCompleted || 0} jobs</span>
                    <span class="text-sm font-bold text-green-400">₱${s.dailyCommission.toFixed(2)}</span>
                </div>
            </div>
        `).join('');

        let recentCommissionHtml = recentCommissions.map(c => `
            <div class="commission-item flex justify-between items-center py-1.5 border-b border-gray-700/30 text-xs">
                <div>
                    <span class="text-gray-400">${new Date(c.time).toLocaleTimeString()}</span>
                    <span class="text-cyan-300 ml-2">${c.customer}</span>
                </div>
                <div>
                    <span class="text-gray-400">${c.staff}</span>
                    <span class="text-green-400 font-bold ml-2">+₱${c.amount.toFixed(2)}</span>
                </div>
            </div>
        `).join('');

        return `<div class="flex justify-between items-center mb-4">
            <div class="flex items-center gap-3">
                ${logoHtml}
                <div>
                    <span class="text-sm text-gray-400"><i class="fas fa-clock"></i> <span id="liveClock" class="live-clock">Loading...</span></span>
                    <span class="text-xs text-cyan-400 ml-3">👤 ${appData.currentUser?.name} (${appData.currentUser?.role})</span>
                    <button onclick="logout()" class="text-xs text-red-400 ml-3 hover:text-red-300"><i class="fas fa-sign-out-alt"></i> Logout</button>
                </div>
            </div>
            <div class="flex items-center gap-3">
                <span id="firebaseStatus" class="firebase-status ${firebaseStatus === 'connected' ? 'firebase-connected' : 'firebase-disconnected'}">
                    ${firebaseStatus === 'connected' ? '<i class="fas fa-check-circle"></i> Firebase Connected' : '<i class="fas fa-cloud"></i> Firebase Offline'}
                </span>
                <div class="text-sm text-gray-400">${appData.systemSettings.businessHours}</div>
            </div>
        </div>
        
        <div class="grid lg:grid-cols-4 gap-4 mb-6">
            <div class="card-premium p-4"><div class="text-gray-400 text-sm">Today's Sales</div><div class="text-2xl font-bold text-green-400">₱${todaySales.toFixed(2)}</div></div>
            <div class="card-premium p-4"><div class="text-gray-400 text-sm">Weekly Sales</div><div class="text-2xl font-bold text-cyan-400">₱${weeklySales.toFixed(2)}</div></div>
            <div class="card-premium p-4"><div class="text-gray-400 text-sm">Monthly Sales</div><div class="text-2xl font-bold text-yellow-400">₱${monthlySales.toFixed(2)}</div></div>
            <div class="card-premium p-4"><div class="text-gray-400 text-sm">Open Services</div><div class="text-2xl font-bold text-orange-400">${openCount}</div></div>
        </div>
        
        <div class="grid lg:grid-cols-4 gap-4 mb-6">
            <div class="card-premium p-4"><div class="text-gray-400 text-sm">Low Stock Items</div><div class="text-2xl font-bold text-red-400">${lowStock}</div></div>
            <div class="card-premium p-4"><div class="text-gray-400 text-sm">Active Staff</div><div class="text-2xl font-bold">${appData.staff.filter(s=>s.isActive).length}</div></div>
            <div class="card-premium p-4"><div class="text-gray-400 text-sm">Total Customers</div><div class="text-2xl font-bold">${appData.customers.length}</div></div>
            <div class="card-premium p-4"><div class="text-gray-400 text-sm">Total Transactions</div><div class="text-2xl font-bold">${appData.transactions.length}</div></div>
        </div>
        
        <div class="grid lg:grid-cols-3 gap-6 mb-6">
            <div class="lg:col-span-2 card-premium p-5">
                <h3 class="text-lg font-bold mb-3"><i class="fas fa-chart-bar"></i> Popular Services</h3>
                <div class="max-h-64 overflow-y-auto">${barHtml || '<div class="text-gray-400 text-center py-4">No service data yet</div>'}</div>
            </div>
            <div class="card-premium p-5">
                <h3 class="text-lg font-bold mb-3 flex items-center gap-2">
                    <i class="fas fa-coins text-yellow-400"></i> Daily Commissions
                    <span class="text-xs text-gray-400 ml-auto">${new Date().toLocaleDateString()}</span>
                </h3>
                <div class="mb-3 pb-2 border-b border-gray-700">
                    <div class="flex justify-between items-center">
                        <span class="text-sm text-gray-400">Total Today</span>
                        <span class="text-2xl font-bold text-green-400">₱${totalDailyCommission.toFixed(2)}</span>
                    </div>
                </div>
                <div class="max-h-48 overflow-y-auto">
                    ${commissionHtml || '<div class="text-gray-400 text-center py-4 text-sm">No commissions yet today</div>'}
                </div>
                <div class="mt-3 pt-2 border-t border-gray-700">
                    <div class="text-xs text-gray-400 flex items-center gap-1">
                        <span class="live-dot"></span> Live updates
                    </div>
                </div>
            </div>
        </div>
        
        <div class="grid lg:grid-cols-2 gap-6">
            <div class="card-premium p-5">
                <h3 class="text-lg font-bold mb-3"><i class="fas fa-chart-line"></i> Sales Trend</h3>
                <canvas id="salesChart" height="200"></canvas>
            </div>
            <div class="card-premium p-5">
                <h3 class="text-lg font-bold mb-3 flex items-center gap-2">
                    <i class="fas fa-bolt text-cyan-400"></i> Recent Commission Activity
                </h3>
                <div class="max-h-48 overflow-y-auto">
                    ${recentCommissionHtml || '<div class="text-gray-400 text-center py-4 text-sm">No recent commission activity</div>'}
                </div>
                <div class="mt-3 pt-2 border-t border-gray-700">
                    <div class="text-xs text-gray-400 flex items-center gap-1">
                        <span class="live-dot"></span> Real-time feed
                    </div>
                </div>
            </div>
        </div>`;
    }

    // ======================== POS ========================
    function renderPOS() {
        if(showLogin) return renderLogin();
        if(!hasRole(['Owner', 'Manager', 'Cashier'])) {
            return `<div class="card-premium p-8 text-center"><i class="fas fa-lock text-4xl text-red-400 mb-4"></i><p class="text-xl">Access Denied. Cashier or higher role required.</p></div>`;
        }
        
        const activeServices = appData.services.filter(s => s.isActive !== false);
        const serviceCategories = [...new Set(activeServices.map(s => s.category))];
        
        let servicesHtml = '';
        serviceCategories.forEach(cat => {
            const catServices = activeServices.filter(s => s.category === cat);
            if(catServices.length) {
                servicesHtml += `<div class="mb-6"><h3 class="text-lg font-bold text-cyan-400 mb-2"><span class="category-badge">${cat}</span></h3><div class="grid grid-cols-2 md:grid-cols-3 gap-2">`;
                catServices.forEach(service => {
                    servicesHtml += `<div class="service-btn bg-gray-900 p-3 rounded-lg" onclick="showSizeModal(${service.id})">
                        <div class="font-bold text-sm">${service.name}</div>
                        <div class="text-xs text-gray-400">${service.desc.substring(0, 35)}</div>
                        <div class="text-cyan-300 text-xs mt-1">From ₱${Object.values(service.prices)[0]}</div>
                    </div>`;
                });
                servicesHtml += `</div></div>`;
            }
        });

        const cartHtml = posCart.map((item, idx) => `<div class="flex justify-between items-center border-b border-gray-700 py-3"><div class="flex-1"><div class="font-semibold">${item.serviceName} (${item.variant})</div><div class="text-xs text-cyan-400">💰 ${item.assignedStaff?.map(s => `${s.name}: ₱${s.splitCommission.toFixed(2)}`).join(', ') || 'No staff'}</div><div class="text-xs text-gray-400">📝 ${item.remarks || 'None'}</div><div>₱${item.price} × ${item.quantity} = <strong>₱${item.total}</strong></div></div><div><input type="number" min="1" value="${item.quantity}" class="w-16 bg-gray-700 rounded text-sm p-1" onchange="updateCartQuantity(${idx}, this.value)"><div class="mt-2"><i class="fas fa-pen text-yellow-400 cursor-pointer mr-2" onclick="addItemRemark(${idx})"></i><i class="fas fa-users text-cyan-400 cursor-pointer mr-2" onclick="reassignStaff(${idx}, '${item.serviceName}', '${item.variant}')"></i><i class="fas fa-trash text-red-500 cursor-pointer" onclick="removeCartItem(${idx})"></i></div></div></div>`).join('');
        
        return `<div class="grid lg:grid-cols-3 gap-6">
            <div class="lg:col-span-2 card-premium p-5 max-h-[75vh] overflow-y-auto">
                <h2 class="text-2xl font-bold mb-4 neon-glow p-2 rounded">Service Menu</h2>
                ${servicesHtml}
            </div>
            <div class="card-premium p-5">
                <h3 class="text-xl font-bold mb-3">Shopping Cart (${posCart.length})</h3>
                <div>
                    <input placeholder="Plate Number" id="plateInput" class="w-full bg-gray-800 p-2 rounded my-1" value="${vehicleInfo.plate}">
                    <input placeholder="Vehicle Model" id="modelInput" class="w-full bg-gray-800 p-2 rounded my-1" value="${vehicleInfo.model}">
                    <button onclick="selectCustomerPopup()" class="bg-blue-600 p-2 rounded w-full">${selectedCustomer ? selectedCustomer.name : "Select Customer"}</button>
                </div>
                <div class="max-h-52 overflow-y-auto my-3">${cartHtml || "<div class='text-gray-500 text-center py-4'>Cart is empty</div>"}</div>
                <div><textarea placeholder="Overall Remarks..." id="cartRemarksInput" class="w-full bg-gray-800 p-2 rounded my-1 text-sm" rows="2">${cartRemarks}</textarea></div>
                <div>Subtotal: ₱${calcSubtotal()}</div>
                <div>Discount: <input id="discountInput" value="${discount}" class="w-24 bg-gray-700 rounded px-2" step="10"></div>
                <div>Loyalty: <input id="loyaltyUse" value="${loyaltyPointsUsed}" class="w-24 bg-gray-700 rounded"></div>
                <div class="font-bold text-xl text-cyan-400">Total: ₱${calcTotal()}</div>
                <select id="paymentMethod" class="bg-gray-800 p-2 w-full rounded my-2">
                    ${appData.systemSettings.defaultPaymentMethods.map(m => `<option>${m}</option>`).join('')}
                </select>
                <input placeholder="Cash Tendered" id="cashTenderedInput" class="bg-gray-800 p-2 w-full rounded" type="number">
                <button onclick="completeSaleWrapper()" class="btn-neon w-full p-3 rounded-xl mt-2">Complete Sale</button>
                <button onclick="saveAsOpenService()" class="bg-orange-600 w-full p-2 rounded-xl mt-2"><i class="fas fa-clock"></i> Save as Open Service</button>
            </div>
        </div>`;
    }

    window.saveAsOpenService = () => {
        if(posCart.length === 0) { alert("Please add services to cart!"); return; }
        const customer = selectedCustomer ? selectedCustomer.name : prompt("Customer name:") || "Walk-in Customer";
        const vehicle = vehicleInfo.model || prompt("Vehicle model:") || "N/A";
        const plate = vehicleInfo.plate || prompt("Plate number:") || "N/A";
        const total = calcTotal();
        const staff = posCart.flatMap(i => i.assignedStaff || []).filter((v, i, a) => a.indexOf(v) === i);
        
        createOpenService(customer, vehicle, plate, posCart, staff, total, cartRemarks);
        alert(`✅ Service saved as Open Service!\nTotal: ₱${total.toFixed(2)}`);
        
        posCart = [];
        selectedCustomer = null;
        vehicleInfo = { plate: "", model: "", size: "Medium" };
        discount = 0;
        loyaltyPointsUsed = 0;
        cartRemarks = "";
        renderApp();
    };

    // ======================== POS HELPERS ========================
    function showSizeModal(serviceId) {
        const service = appData.services.find(s => s.id === serviceId);
        if(!service.isActive) { alert("This service is currently inactive."); return; }
        let sizeHtml = service.sizes.map(size => `<div class="size-option bg-gray-800 p-3 rounded-lg cursor-pointer hover:bg-cyan-500 hover:text-black transition-all m-2 inline-block" data-size="${size}" data-price="${service.prices[size]}"><span class="font-bold">${size}</span><br><span class="text-cyan-300">₱${service.prices[size]}</span></div>`).join('');
        let modalHtml = `<div id="sizeModal" class="modal"><div class="modal-content"><h3 class="text-xl font-bold mb-2">${service.name}</h3><p class="text-sm text-gray-400 mb-4">${service.desc}</p><p class="text-cyan-400 mb-3">Select Vehicle Size:</p><div id="sizeOptions" class="flex flex-wrap gap-2 mb-4 justify-center">${sizeHtml}</div><div><label class="text-sm text-gray-400">Item Remarks:</label><input type="text" id="itemRemarkInput" class="w-full bg-gray-800 p-2 rounded mt-1" placeholder="Special instructions..."></div><div class="flex gap-3 mt-4"><button onclick="confirmAddToCartWithSize(${serviceId})" class="btn-neon flex-1 p-2 rounded">Add to Cart</button><button onclick="closeSizeModal()" class="bg-gray-700 flex-1 p-2 rounded">Cancel</button></div></div></div>`;
        document.body.insertAdjacentHTML('beforeend', modalHtml);
        document.querySelectorAll('.size-option').forEach(opt => { opt.addEventListener('click', function() { document.querySelectorAll('.size-option').forEach(o => o.classList.remove('bg-cyan-500', 'text-black')); document.querySelectorAll('.size-option').forEach(o => o.classList.add('bg-gray-800')); this.classList.remove('bg-gray-800'); this.classList.add('bg-cyan-500', 'text-black'); this.dataset.selected = "true"; }); });
    }

    window.confirmAddToCartWithSize = (serviceId) => {
        const selectedSizeElem = document.querySelector('.size-option[data-selected="true"]') || document.querySelector('.size-option');
        if(!selectedSizeElem) { alert("Please select a size first!"); return; }
        const size = selectedSizeElem.dataset.size;
        const price = parseFloat(selectedSizeElem.dataset.price);
        const service = appData.services.find(s => s.id === serviceId);
        const remark = document.getElementById('itemRemarkInput')?.value || "";
        posCart.push({ serviceId, serviceName: service.name, variant: size, price, quantity: 1, total: price, assignedStaff: [], remarks: remark });
        closeSizeModal();
        showStaffAssignmentModal(posCart.length - 1, service.name, size);
        renderApp();
    };

    function showStaffAssignmentModal(cartIndex, serviceName, size) {
        const activeStaff = appData.staff.filter(s => s.isActive);
        let staffOptionsHtml = activeStaff.map(s => { 
            const fullCommission = getCommissionAmount(serviceName, size, s.commissionTier);
            return `<label class="flex items-center justify-between p-3 bg-gray-800 rounded cursor-pointer hover:bg-gray-700 mb-2"><div class="flex items-center space-x-3"><input type="checkbox" value="${s.id}" class="staff-checkbox w-5 h-5" data-name="${s.name}" data-tier="${s.commissionTier}" data-full-comm="${fullCommission}"><span><strong>${s.name}</strong> - ${s.position}</span></div><div class="text-right"><span class="commission-badge">${s.commissionTier}%</span><div class="text-xs text-green-400">₱${fullCommission}</div></div></label>`; 
        }).join('');
        let modalHtml = `<div id="staffModal" class="modal"><div class="modal-content"><h3 class="text-xl font-bold mb-2">Assign Staff: ${serviceName} (${size})</h3><p class="text-sm text-yellow-400 mb-2">Commission Split Rule: ${appData.systemSettings.commissionSplitRule}</p><div id="staffCheckboxes" class="space-y-2 mb-4">${staffOptionsHtml}</div><div id="commissionPreview" class="bg-gray-900 p-3 rounded-lg mb-3 text-sm"><span id="previewText">Select staff to see split breakdown</span></div><div class="flex gap-3"><button onclick="confirmStaffAssignment(${cartIndex}, '${serviceName}', '${size}')" class="btn-neon flex-1 p-2 rounded">Confirm</button><button onclick="closeStaffModal()" class="bg-gray-700 flex-1 p-2 rounded">Skip</button></div></div></div>`;
        document.body.insertAdjacentHTML('beforeend', modalHtml);
        document.querySelectorAll('.staff-checkbox').forEach(cb => cb.addEventListener('change', updateSplitPreview));
        function updateSplitPreview() { const selected = []; document.querySelectorAll('.staff-checkbox:checked').forEach(cb => selected.push({ name: cb.dataset.name, tier: parseInt(cb.dataset.tier), fullCommission: parseFloat(cb.dataset.full_comm) })); if(selected.length === 1) document.getElementById('previewText').innerHTML = `<div class="text-green-400">💰 1 Staff - Full Commission</div><div>${selected[0].name}: ₱${selected[0].fullCommission}</div>`; else if(selected.length === 2) document.getElementById('previewText').innerHTML = `<div class="text-green-400">👥 2 Staff - EQUAL SPLIT (50/50)</div><div>${selected[0].name}: ₱${(selected[0].fullCommission/2).toFixed(2)} | ${selected[1].name}: ₱${(selected[1].fullCommission/2).toFixed(2)}</div><div class="text-cyan-400">Total: ₱${(selected[0].fullCommission/2 + selected[1].fullCommission/2).toFixed(2)}</div>`; else document.getElementById('previewText').innerHTML = '<span class="text-gray-400">Select 1 or 2 staff members</span>'; }
    }

    window.confirmStaffAssignment = (cartIndex, serviceName, size) => {
        const selected = []; document.querySelectorAll('.staff-checkbox:checked').forEach(cb => selected.push({ id: parseInt(cb.value), name: cb.dataset.name, commissionTier: parseInt(cb.dataset.tier), fullCommission: getCommissionAmount(serviceName, size, parseInt(cb.dataset.tier)) }));
        if(selected.length > 0) { const assignedWithSplit = selected.map(s => ({ ...s, splitCommission: selected.length === 2 ? s.fullCommission / 2 : s.fullCommission })); posCart[cartIndex].assignedStaff = assignedWithSplit; }
        closeStaffModal(); renderApp();
    };

    window.closeStaffModal = () => document.getElementById('staffModal')?.remove();
    window.closeSizeModal = () => document.getElementById('sizeModal')?.remove();
    window.addItemRemark = (idx) => { const remark = prompt("Enter remark for this service:", posCart[idx].remarks || ""); if(remark !== null) { posCart[idx].remarks = remark; renderApp(); } };
    window.reassignStaff = (idx, name, size) => showStaffAssignmentModal(idx, name, size);

    function removeCartItem(idx) { posCart.splice(idx,1); renderApp(); }
    function updateCartQuantity(idx, qty) { if(qty>0) { posCart[idx].quantity = qty; posCart[idx].total = posCart[idx].price * qty; renderApp(); } }
    function calcSubtotal() { return posCart.reduce((sum,item)=> sum + item.total,0); }
    function calcTotal() { let total = calcSubtotal() - discount - loyaltyPointsUsed; return total > 0 ? total : 0; }

    function completeSaleWrapper() {
        selectedPayment = document.getElementById('paymentMethod')?.value || "Cash";
        cashTendered = parseFloat(document.getElementById('cashTenderedInput')?.value || 0);
        discount = parseFloat(document.getElementById('discountInput')?.value || 0);
        loyaltyPointsUsed = parseFloat(document.getElementById('loyaltyUse')?.value || 0);
        cartRemarks = document.getElementById('cartRemarksInput')?.value || "";
        completeSale();
    }

    function completeSale() {
        if(posCart.length === 0) { alert("Please add services to cart!"); return; }
        const total = calcTotal();
        if(selectedPayment === "Cash" && cashTendered < total) { alert("Insufficient cash!"); return; }
        
        const staffCommissionMap = new Map();
        posCart.forEach(item => { if(item.assignedStaff) item.assignedStaff.forEach(staff => { const current = staffCommissionMap.get(staff.id) || { commission: 0, jobs: 0, name: staff.name, tier: staff.commissionTier }; current.commission += staff.splitCommission; current.jobs += 1; staffCommissionMap.set(staff.id, current); }); });
        
        const transaction = { 
            id: generateReceiptNo(), 
            date: new Date().toISOString(), 
            customer: selectedCustomer ? selectedCustomer.name : "Walk-in Customer", 
            customerId: selectedCustomer?.id, 
            vehicle: vehicleInfo.model || "N/A", 
            plate: vehicleInfo.plate || "N/A", 
            items: posCart.map(i => ({ ...i })), 
            staffCommissions: Object.fromEntries([...staffCommissionMap.entries()]), 
            subtotal: calcSubtotal(), 
            discount, 
            loyaltyUsed: loyaltyPointsUsed, 
            total, 
            paymentMethod: selectedPayment, 
            cashTendered: selectedPayment === "Cash" ? cashTendered : total, 
            change: selectedPayment === "Cash" ? (cashTendered - total) : 0,
            remarks: cartRemarks,
            itemRemarks: posCart.map(i => ({ service: i.serviceName, remark: i.remarks || 'None' }))
        };
        
        appData.transactions.unshift(transaction);
        staffCommissionMap.forEach((data, staffId) => { const staff = appData.staff.find(s => s.id === staffId); if(staff) { staff.totalCommission = (staff.totalCommission || 0) + data.commission; staff.jobsCompleted += data.jobs; staff.salesGenerated += transaction.items.filter(i => i.assignedStaff?.some(s => s.id === staffId)).reduce((sum, i) => sum + i.total, 0); } });
        
        if(selectedCustomer) { let earnedPoints = Math.floor(total * appData.systemSettings.loyaltyPointsPerPeso); selectedCustomer.loyaltyPoints = (selectedCustomer.loyaltyPoints || 0) + earnedPoints; const idx = appData.customers.findIndex(c=>c.id === selectedCustomer.id); if(idx !== -1) appData.customers[idx].loyaltyPoints = selectedCustomer.loyaltyPoints; }
        
        saveToLocal();
        alert(`✅ SALE COMPLETED!\nTotal: ₱${total.toFixed(2)}`);
        printReceipt(transaction);
        posCart = []; selectedCustomer = null; vehicleInfo = { plate: "", model: "", size: "Medium" }; discount = 0; loyaltyPointsUsed = 0; cashTendered = 0; cartRemarks = "";
        renderApp();
    }

    function printReceipt(tx) {
        const win = window.open('', '_blank');
        const logoHtml = appData.systemSettings.companyLogo ? 
            `<div class="center"><img src="${appData.systemSettings.companyLogo}" style="max-width:80px;max-height:60px;object-fit:contain;"></div>` : '';
        win.document.write(`<html><head><title>Receipt</title><style>
            body{font-family:monospace;width:80mm;margin:0 auto;padding:10px;background:white;color:black;}
            .center{text-align:center;}.line{border-top:1px dashed #000;margin:8px 0;}.bold{font-weight:bold;}.small{font-size:10px;}
            img{max-width:80px;max-height:60px;object-fit:contain;}
        </style></head><body>
            ${logoHtml}
            <div class="center"><strong>${appData.systemSettings.companyName}</strong><br>${appData.systemSettings.tagline}<br>${appData.systemSettings.address}<br>Tel: ${appData.systemSettings.phone}</div>
            <div class="line"></div>
            <div>Receipt: ${tx.id}<br>Date: ${new Date(tx.date).toLocaleString()}<br>Customer: ${tx.customer}<br>Vehicle: ${tx.vehicle} | ${tx.plate}</div>
            <div class="line"></div>
            ${tx.items.map(i => `<div><strong>${i.serviceName}</strong> (${i.variant}) x${i.quantity}<br>₱${(i.price*i.quantity).toFixed(2)}${i.remarks ? `<br><span class="small">📝 ${i.remarks}</span>` : ''}<br><span class="small">Staff: ${i.assignedStaff?.map(s=>`${s.name} (₱${s.splitCommission.toFixed(2)})`).join(', ') || 'Unassigned'}</span></div><div class="line"></div>`).join('')}
            ${tx.remarks ? `<div class="small">📝 Overall: ${tx.remarks}</div><div class="line"></div>` : ''}
            <div>Subtotal: ₱${tx.subtotal.toFixed(2)}<br>Discount: ₱${tx.discount.toFixed(2)}<br><span class="bold">TOTAL: ₱${tx.total.toFixed(2)}</span><br>Payment: ${tx.paymentMethod}<br>Change: ₱${tx.change.toFixed(2)}</div>
            <div class="line"></div>
            <div class="center">Thank You!<br>${appData.systemSettings.receiptFooter}</div>
        </body></html>`);
        win.document.close(); win.print();
    }

    // ======================== CUSTOMER FUNCTIONS ========================
    function renderCustomers() {
        if(showLogin) return renderLogin();
        return `<div class="card-premium p-6"><div class="flex justify-between mb-4"><h2 class="text-2xl font-bold">👥 Customer Database</h2><button onclick="addCustomerPopup()" class="bg-green-600 px-3 py-1 rounded">+ Add</button></div>
            <div class="overflow-x-auto max-h-96 overflow-y-auto"><table class="w-full"><thead class="sticky top-0 bg-gray-900 border-b border-cyan-400"><tr><th>Name</th><th>Contact</th><th>Vehicle</th><th>Plate</th><th>Points</th><th>History</th></tr></thead><tbody>
            ${appData.customers.map(c => `<tr class="border-b border-gray-800"><td class="py-2">${c.name}</td><td>${c.mobile || "—"}</td><td>${c.vehicleInfo || "—"}</td><td>${c.plate || "—"}</td><td>${c.loyaltyPoints}</td><td><button onclick="viewCustomerHistory(${c.id})" class="text-cyan-400 text-sm">View</button></td></tr>`).join('')}</tbody></table></div></div>`;
    }

    window.viewCustomerHistory = (id) => {
        const tx = appData.transactions.filter(t => t.customerId === id);
        alert(`Customer has ${tx.length} transaction(s). Total spent: ₱${tx.reduce((s,t)=>s+t.total,0).toFixed(2)}`);
    };

    window.addCustomerPopup = () => { let name = prompt("Customer name:"); if(name) { appData.customers.push({ id: Date.now(), name, mobile: "", loyaltyPoints: 0 }); saveToLocal(); renderApp(); } };
    window.selectCustomerPopup = () => { let name = prompt("Customer name:"); if(name) { let existing = appData.customers.find(c=>c.name.toLowerCase()===name.toLowerCase()); if(existing) selectedCustomer = existing; else { let newCust = { id: Date.now(), name, loyaltyPoints:0 }; appData.customers.push(newCust); selectedCustomer = newCust; } renderApp(); } };

    // ======================== ADMIN DASHBOARD ========================
    function renderAdminDashboard() {
        if(showLogin) return renderLogin();
        if(!hasRole(['Owner', 'Manager'])) {
            return `<div class="card-premium p-8 text-center"><i class="fas fa-lock text-4xl text-red-400 mb-4"></i><p class="text-xl">Access Denied. Manager or Owner role required.</p></div>`;
        }
        
        const adminPanels = [
            { id: 'services', icon: 'fa-concierge-bell', label: 'Services' },
            { id: 'categories', icon: 'fa-tags', label: 'Categories' },
            { id: 'staff', icon: 'fa-users', label: 'Staff' },
            { id: 'inventory', icon: 'fa-boxes', label: 'Inventory' },
            { id: 'customers', icon: 'fa-user-friends', label: 'Customers' },
            { id: 'expenses', icon: 'fa-money-bill-wave', label: 'Expenses' },
            { id: 'transactions', icon: 'fa-receipt', label: 'Transactions' },
            { id: 'commissions', icon: 'fa-percent', label: 'Commission Matrix' },
            { id: 'users', icon: 'fa-user-shield', label: 'Users' },
            { id: 'system', icon: 'fa-cogs', label: 'System Settings' }
        ];
        
        let content = '';
        if(currentAdminPanel === 'services') content = renderAdminServices();
        else if(currentAdminPanel === 'categories') content = renderAdminCategories();
        else if(currentAdminPanel === 'staff') content = renderStaffDashboard();
        else if(currentAdminPanel === 'inventory') content = renderInventory();
        else if(currentAdminPanel === 'customers') content = renderCustomers();
        else if(currentAdminPanel === 'expenses') content = renderAdminExpenses();
        else if(currentAdminPanel === 'transactions') content = renderAdminTransactions();
        else if(currentAdminPanel === 'commissions') content = renderAdminCommissions();
        else if(currentAdminPanel === 'users') content = renderAdminUsers();
        else if(currentAdminPanel === 'system') content = renderAdminSystem();
        
        return `<div class="admin-grid">
            <div class="admin-sidebar">
                <h3 class="text-lg font-bold mb-4 neon-glow p-2 rounded text-center"><i class="fas fa-crown"></i> Admin Panel</h3>
                ${adminPanels.map(p => `<div class="admin-sidebar-item ${currentAdminPanel === p.id ? 'active' : ''}" onclick="navigateAdmin('${p.id}')"><i class="fas ${p.icon}"></i> ${p.label}</div>`).join('')}
                <div class="mt-6 p-3 bg-gray-800 rounded-lg text-xs">
                    <div class="text-gray-400">System Status</div>
                    <div class="grid grid-cols-2 gap-1 mt-2">
                        <div>Services: ${appData.services.length}</div>
                        <div>Open Services: ${appData.openServices.filter(s=>s.status!=='completed').length}</div>
                        <div>Staff: ${appData.staff.length}</div>
                        <div>TX: ${appData.transactions.length}</div>
                    </div>
                    <div class="mt-2 text-cyan-400 text-[10px]">Last Saved: ${new Date().toLocaleTimeString()}</div>
                    <div class="mt-1 text-gray-500 text-[10px]">👤 ${appData.currentUser?.name} (${appData.currentUser?.role})</div>
                </div>
            </div>
            <div class="admin-content card-premium p-5">${content}</div>
        </div>`;
    }

    // ======================== ADMIN - SERVICES ========================
    function renderAdminServices() {
        if(showLogin) return renderLogin();
        return `<h3 class="text-xl font-bold mb-4"><i class="fas fa-concierge-bell"></i> Services Management</h3>
            <button onclick="showAddServiceModal()" class="bg-green-600 px-3 py-1 rounded mb-4">+ Add Service</button>
            <div class="overflow-x-auto max-h-96 overflow-y-auto"><table class="w-full text-sm"><thead class="sticky top-0 bg-gray-900 border-b border-cyan-400"><tr><th>Name</th><th>Category</th><th>Sizes</th><th>Prices</th><th>Status</th><th>Actions</th></tr></thead><tbody>
            ${appData.services.map(s => `<tr class="border-b border-gray-700"><td class="py-2 editable-field cursor-pointer" onclick="editServiceName(${s.id})">${s.name}</td><td><span class="category-badge">${s.category}</span></td><td class="text-xs">${s.sizes.join(', ')}</td><td class="text-xs">${Object.entries(s.prices).map(([k,v])=>`${k}:₱${v}`).join(' ')}</td><td><button onclick="toggleServiceStatus(${s.id})" class="text-${s.isActive !== false ? 'green' : 'red'}-400">${s.isActive !== false ? 'Active' : 'Inactive'}</button></td><td><button onclick="editServicePrice(${s.id})" class="text-cyan-400 mr-2"><i class="fas fa-edit"></i></button><button onclick="deleteServiceItem(${s.id})" class="text-red-400"><i class="fas fa-trash"></i></button></td></tr>`).join('')}</tbody></table></div>`;
    }

    window.showAddServiceModal = () => {
        const categoryOptions = appData.categories.map(c => `<option>${c}</option>`).join('');
        let modalHtml = `<div id="adminModal" class="modal"><div class="modal-content"><h3 class="text-xl font-bold mb-4">Add New Service</h3><div class="space-y-3"><div><label class="text-gray-400">Service Name</label><input type="text" id="newServiceName" class="w-full bg-gray-800 p-2 rounded"></div><div><label class="text-gray-400">Description</label><input type="text" id="newServiceDesc" class="w-full bg-gray-800 p-2 rounded"></div><div><label class="text-gray-400">Category</label><select id="newServiceCat" class="w-full bg-gray-800 p-2 rounded">${categoryOptions}</select></div><div><label class="text-gray-400">Sizes (comma separated)</label><input type="text" id="newServiceSizes" class="w-full bg-gray-800 p-2 rounded" placeholder="Small, Medium, Large"></div><div><label class="text-gray-400">Prices (comma separated matching sizes)</label><input type="text" id="newServicePrices" class="w-full bg-gray-800 p-2 rounded" placeholder="200, 230, 250"></div><div class="flex gap-3"><button onclick="confirmAddService()" class="btn-neon flex-1 p-2 rounded">Create</button><button onclick="closeAdminModal()" class="bg-gray-700 flex-1 p-2 rounded">Cancel</button></div></div></div></div>`;
        document.body.insertAdjacentHTML('beforeend', modalHtml);
    };
    window.confirmAddService = () => {
        const name = document.getElementById('newServiceName')?.value; if(!name) { alert("Name required!"); return; }
        const desc = document.getElementById('newServiceDesc')?.value || "";
        const category = document.getElementById('newServiceCat')?.value || "Specialty";
        const sizes = document.getElementById('newServiceSizes')?.value.split(',').map(s => s.trim());
        const pricesArr = document.getElementById('newServicePrices')?.value.split(',').map(p => parseFloat(p.trim()));
        const prices = {}; sizes.forEach((size, idx) => { if(pricesArr[idx]) prices[size] = pricesArr[idx]; });
        const newService = { id: Date.now(), name: name.toUpperCase(), desc, sizes, prices, category, isActive: true };
        appData.services.push(newService);
        sizes.forEach(size => {
            const price = prices[size];
            if (price) {
                const tier30 = Math.round(price * 0.26);
                const tier35 = Math.round(price * 0.30);
                const tier40 = Math.round(price * 0.35);
                const key = `${name.toUpperCase()}_${size}`;
                appData.commissionMatrix[key] = [tier30, tier35, tier40];
            }
        });
        saveToLocal(); closeAdminModal(); renderApp(); alert(`✅ Service "${name}" added!`);
    };
    window.editServiceName = (id) => { let s = appData.services.find(s=>s.id===id); if(s) { let newName = prompt("New name:", s.name); if(newName) { const oldName = s.name; s.name = newName.toUpperCase(); 
                const newMatrix = {};
                Object.keys(appData.commissionMatrix).forEach(key => {
                    if (key.startsWith(oldName)) {
                        const newKey = key.replace(oldName, s.name);
                        newMatrix[newKey] = appData.commissionMatrix[key];
                    } else {
                        newMatrix[key] = appData.commissionMatrix[key];
                    }
                });
                appData.commissionMatrix = newMatrix;
                saveToLocal(); renderApp(); } } };
    window.editServicePrice = (id) => { let s = appData.services.find(s=>s.id===id); if(s) { let size = prompt("Size to edit:", s.sizes[0]); let price = parseFloat(prompt("New price:")); if(size && price) { s.prices[size] = price; 
                const key = `${s.name}_${size}`;
                const tier30 = Math.round(price * 0.26);
                const tier35 = Math.round(price * 0.30);
                const tier40 = Math.round(price * 0.35);
                appData.commissionMatrix[key] = [tier30, tier35, tier40];
                saveToLocal(); renderApp(); } } };
    window.toggleServiceStatus = (id) => { let s = appData.services.find(s=>s.id===id); if(s) { s.isActive = !s.isActive; saveToLocal(); renderApp(); } };
    window.deleteServiceItem = (id) => { if(confirm("Delete this service?")) { const service = appData.services.find(s => s.id === id); if(service) { 
                Object.keys(appData.commissionMatrix).forEach(key => {
                    if (key.startsWith(service.name)) {
                        delete appData.commissionMatrix[key];
                    }
                });
                appData.services = appData.services.filter(s => s.id !== id);
                saveToLocal(); renderApp(); } } };

    // ======================== ADMIN - CATEGORIES ========================
    function renderAdminCategories() {
        if(showLogin) return renderLogin();
        return `<h3 class="text-xl font-bold mb-4"><i class="fas fa-tags"></i> Categories Management</h3>
            <button onclick="showAddCategoryModal()" class="bg-green-600 px-3 py-1 rounded mb-4">+ Add Category</button>
            <div class="overflow-x-auto">
                <table class="w-full text-sm">
                    <thead class="border-b border-cyan-400">
                        <tr><th>Category Name</th><th>Services Count</th><th>Actions</th></tr>
                    </thead>
                    <tbody>
                        ${appData.categories.map(cat => {
                            const count = appData.services.filter(s => s.category === cat).length;
                            return `<tr class="border-b border-gray-700">
                                <td class="py-2"><span class="category-badge">${cat}</span></td>
                                <td>${count} service(s)</td>
                                <td>
                                    <button onclick="editCategory('${cat}')" class="text-cyan-400 mr-2"><i class="fas fa-edit"></i></button>
                                    <button onclick="deleteCategory('${cat}')" class="text-red-400"><i class="fas fa-trash"></i></button>
                                </td>
                            </tr>`;
                        }).join('')}
                    </tbody>
                </table>
            </div>`;
    }

    window.showAddCategoryModal = () => {
        let modalHtml = `<div id="adminModal" class="modal"><div class="modal-content"><h3 class="text-xl font-bold mb-4">Add New Category</h3><div class="space-y-3"><div><label class="text-gray-400">Category Name</label><input type="text" id="newCategoryName" class="w-full bg-gray-800 p-2 rounded" placeholder="e.g., Interior Services"></div><div class="flex gap-3 mt-4"><button onclick="confirmAddCategory()" class="btn-neon flex-1 p-2 rounded">Add Category</button><button onclick="closeAdminModal()" class="bg-gray-700 flex-1 p-2 rounded">Cancel</button></div></div></div></div>`;
        document.body.insertAdjacentHTML('beforeend', modalHtml);
    };
    window.confirmAddCategory = () => {
        const name = document.getElementById('newCategoryName')?.value?.trim();
        if(!name) { alert("Category name required!"); return; }
        if(appData.categories.includes(name)) { alert("Category already exists!"); return; }
        appData.categories.push(name);
        saveToLocal(); closeAdminModal(); renderApp(); alert(`Category "${name}" added!`);
    };
    window.editCategory = (oldName) => {
        const newName = prompt("Edit category name:", oldName);
        if(newName && newName.trim() && newName !== oldName) {
            if(appData.categories.includes(newName.trim())) { alert("Category already exists!"); return; }
            const idx = appData.categories.indexOf(oldName);
            if(idx !== -1) {
                appData.categories[idx] = newName.trim();
                appData.services.forEach(s => { if(s.category === oldName) s.category = newName.trim(); });
                saveToLocal(); renderApp(); alert(`Category updated to "${newName}"!`);
            }
        }
    };
    window.deleteCategory = (cat) => {
        if(confirm(`Delete category "${cat}"?`)) {
            appData.categories = appData.categories.filter(c => c !== cat);
            appData.services.forEach(s => { if(s.category === cat) s.category = "Specialty"; });
            saveToLocal(); renderApp();
        }
    };

    // ======================== INVENTORY ========================
    function renderInventory() {
        if(showLogin) return renderLogin();
        if(!hasRole(['Owner', 'Manager'])) {
            return `<div class="card-premium p-8 text-center"><i class="fas fa-lock text-4xl text-red-400 mb-4"></i><p class="text-xl">Access Denied. Manager or Owner role required.</p></div>`;
        }
        return `<div class="card-premium p-6"><div class="flex justify-between mb-4"><h2 class="text-2xl font-bold">📦 Inventory Management</h2><button onclick="addInventoryItem()" class="bg-green-600 px-3 py-1 rounded">+ Add Item</button></div>
            <div class="overflow-x-auto max-h-96 overflow-y-auto"><table class="w-full text-sm"><thead class="sticky top-0 bg-gray-900 border-b border-cyan-400"><tr><th>Item</th><th>Category</th><th>Quantity</th><th>Unit</th><th>Reorder Level</th><th>Status</th><th>Actions</th></tr></thead><tbody>
            ${appData.inventory.map(i => {
                const status = i.quantity <= i.reorderLevel ? 'Low Stock' : i.quantity <= i.reorderLevel * 2 ? 'Medium' : 'High';
                const color = i.quantity <= i.reorderLevel ? 'text-red-400' : i.quantity <= i.reorderLevel * 2 ? 'text-yellow-400' : 'text-green-400';
                return `<tr class="border-b border-gray-700"><td class="py-2">${i.name}</td><td>${i.category}</td><td class="${color} font-bold">${i.quantity}</td><td>${i.unit}</td><td>${i.reorderLevel}</td><td><span class="status-badge ${i.quantity <= i.reorderLevel ? 'bg-red-400/20 text-red-400' : 'bg-green-400/20 text-green-400'}">${status}</span></td><td><button onclick="editInventory(${i.id})" class="text-cyan-400 mr-2"><i class="fas fa-edit"></i></button><button onclick="deleteInventory(${i.id})" class="text-red-400"><i class="fas fa-trash"></i></button></td></tr>`;
            }).join('')}</tbody></table></div></div>`;
    }

    window.addInventoryItem = () => {
        let name = prompt("Item name:"); if(!name) return;
        let category = prompt("Category:") || "General";
        let quantity = parseInt(prompt("Quantity:")) || 0;
        let unit = prompt("Unit:") || "Pieces";
        let reorderLevel = parseInt(prompt("Reorder level:")) || 10;
        appData.inventory.push({ id: Date.now(), name, category, quantity, unit, reorderLevel, price: 0, supplier: "" });
        saveToLocal(); renderApp();
    };
    window.editInventory = (id) => {
        let item = appData.inventory.find(i => i.id === id);
        if(item) {
            item.name = prompt("Item name:", item.name) || item.name;
            item.quantity = parseInt(prompt("Quantity:", item.quantity)) || item.quantity;
            item.reorderLevel = parseInt(prompt("Reorder level:", item.reorderLevel)) || item.reorderLevel;
            saveToLocal(); renderApp();
        }
    };
    window.deleteInventory = (id) => { if(confirm("Delete inventory item?")) { appData.inventory = appData.inventory.filter(i => i.id !== id); saveToLocal(); renderApp(); } };

    // ======================== ADMIN - EXPENSES ========================
    function renderAdminExpenses() {
        if(showLogin) return renderLogin();
        return `<h3 class="text-xl font-bold mb-4"><i class="fas fa-money-bill-wave"></i> Expenses Management</h3>
            <button onclick="showAddExpenseModal()" class="bg-green-600 px-3 py-1 rounded mb-4">+ Add Expense</button>
            <div class="overflow-x-auto max-h-96 overflow-y-auto">
                <table class="w-full text-sm">
                    <thead class="sticky top-0 bg-gray-900 border-b border-cyan-400">
                        <tr><th>Date</th><th>Description</th><th>Category</th><th>Amount</th><th class="text-center">Actions</th></tr>
                    </thead>
                    <tbody>
                        ${appData.expenses.map(e => `
                            <tr class="border-b border-gray-700 hover:bg-gray-800 transition">
                                <td class="py-2"><span class="editable-field cursor-pointer" onclick="editExpenseDate(${e.id})">${e.date}</span></td>
                                <td><span class="editable-field cursor-pointer" onclick="editExpenseDescAdmin(${e.id})">${e.description}</span></td>
                                <td><span class="editable-field cursor-pointer" onclick="editExpenseCategory(${e.id})">${e.category || "General"}</span></td>
                                <td><span class="editable-field cursor-pointer text-red-400" onclick="editExpenseAmtAdmin(${e.id})">₱${e.amount.toFixed(2)}</span></td>
                                <td class="text-center">
                                    <button onclick="editExpenseFull(${e.id})" class="action-btn bg-cyan-600 text-white hover:bg-cyan-700"><i class="fas fa-edit"></i></button>
                                    <button onclick="deleteExpenseItem(${e.id})" class="action-btn bg-red-600 text-white hover:bg-red-700"><i class="fas fa-trash"></i></button>
                                </td>
                            </tr>
                        `).join('')}
                    </tbody>
                    <tfoot class="sticky bottom-0 bg-gray-900 border-t border-cyan-400">
                        <tr><td colspan="3" class="py-2 text-right font-bold">TOTAL:</td><td class="py-2 text-red-400 font-bold">₱${appData.expenses.reduce((s,e)=>s+e.amount,0).toFixed(2)}</td><td></td></tr>
                    </tfoot>
                </table>
            </div>`;
    }

    window.showAddExpenseModal = () => {
        let modalHtml = `<div id="adminModal" class="modal"><div class="modal-content"><h3 class="text-xl font-bold mb-4">Add New Expense</h3><div class="space-y-3"><div><label>Description</label><input type="text" id="expenseDescInput" class="w-full bg-gray-800 p-2 rounded"></div><div><label>Amount</label><input type="number" id="expenseAmtInput" class="w-full bg-gray-800 p-2 rounded"></div><div><label>Category</label><input type="text" id="expenseCatInput" class="w-full bg-gray-800 p-2 rounded" placeholder="Supplies, Utilities"></div><div><label>Date</label><input type="date" id="expenseDateInput" class="w-full bg-gray-800 p-2 rounded" value="${new Date().toISOString().slice(0,10)}"></div><div class="flex gap-3 mt-4"><button onclick="confirmAddExpense()" class="btn-neon flex-1 p-2 rounded">Add</button><button onclick="closeAdminModal()" class="bg-gray-700 flex-1 p-2 rounded">Cancel</button></div></div></div></div>`;
        document.body.insertAdjacentHTML('beforeend', modalHtml);
    };
    window.confirmAddExpense = () => {
        const desc = document.getElementById('expenseDescInput')?.value;
        const amt = parseFloat(document.getElementById('expenseAmtInput')?.value);
        const cat = document.getElementById('expenseCatInput')?.value || "General";
        const date = document.getElementById('expenseDateInput')?.value;
        if(!desc) { alert("Description required!"); return; }
        if(isNaN(amt) || amt <= 0) { alert("Valid amount required!"); return; }
        appData.expenses.push({ id: Date.now(), description: desc, amount: amt, date: date || new Date().toISOString().slice(0,10), category: cat });
        saveToLocal(); closeAdminModal(); renderApp();
    };
    window.editExpenseFull = (id) => {
        const e = appData.expenses.find(ex => ex.id === id);
        if(!e) return;
        let modalHtml = `<div id="adminModal" class="modal"><div class="modal-content"><h3 class="text-xl font-bold mb-4">Edit Expense</h3><div class="space-y-3"><div><label>Description</label><input type="text" id="editExpenseDesc" class="w-full bg-gray-800 p-2 rounded" value="${e.description}"></div><div><label>Amount</label><input type="number" id="editExpenseAmt" class="w-full bg-gray-800 p-2 rounded" value="${e.amount}"></div><div><label>Category</label><input type="text" id="editExpenseCat" class="w-full bg-gray-800 p-2 rounded" value="${e.category || "General"}"></div><div><label>Date</label><input type="date" id="editExpenseDate" class="w-full bg-gray-800 p-2 rounded" value="${e.date}"></div><div class="flex gap-3 mt-4"><button onclick="confirmEditExpense(${id})" class="btn-neon flex-1 p-2 rounded">Update</button><button onclick="closeAdminModal()" class="bg-gray-700 flex-1 p-2 rounded">Cancel</button></div></div></div></div>`;
        document.body.insertAdjacentHTML('beforeend', modalHtml);
    };
    window.confirmEditExpense = (id) => {
        const e = appData.expenses.find(ex => ex.id === id);
        if(!e) return;
        const desc = document.getElementById('editExpenseDesc')?.value;
        const amt = parseFloat(document.getElementById('editExpenseAmt')?.value);
        const cat = document.getElementById('editExpenseCat')?.value;
        const date = document.getElementById('editExpenseDate')?.value;
        if(desc) e.description = desc;
        if(!isNaN(amt) && amt > 0) e.amount = amt;
        if(cat) e.category = cat;
        if(date) e.date = date;
        saveToLocal(); closeAdminModal(); renderApp();
    };
    window.editExpenseDescAdmin = (id) => { const e = appData.expenses.find(ex => ex.id === id); if(e) { const val = prompt("Edit description:", e.description); if(val !== null) { e.description = val; saveToLocal(); renderApp(); } } };
    window.editExpenseAmtAdmin = (id) => { const e = appData.expenses.find(ex => ex.id === id); if(e) { const val = parseFloat(prompt("Edit amount:", e.amount)); if(!isNaN(val) && val > 0) { e.amount = val; saveToLocal(); renderApp(); } } };
    window.editExpenseCategory = (id) => { const e = appData.expenses.find(ex => ex.id === id); if(e) { const val = prompt("Edit category:", e.category || "General"); if(val !== null) { e.category = val; saveToLocal(); renderApp(); } } };
    window.editExpenseDate = (id) => { const e = appData.expenses.find(ex => ex.id === id); if(e) { const val = prompt("Edit date (YYYY-MM-DD):", e.date); if(val !== null) { e.date = val; saveToLocal(); renderApp(); } } };
    window.deleteExpenseItem = (id) => { if(confirm("Delete this expense?")) { appData.expenses = appData.expenses.filter(e => e.id !== id); saveToLocal(); renderApp(); } };

    // ======================== ADMIN - TRANSACTIONS ========================
    function renderAdminTransactions() {
        if(showLogin) return renderLogin();
        const totalIncome = appData.transactions.reduce((s,t)=>s+t.total,0);
        return `<h3 class="text-xl font-bold mb-4"><i class="fas fa-receipt"></i> Transactions Management</h3>
            <button onclick="showAddTransactionModal()" class="bg-green-600 px-3 py-1 rounded mb-4">+ Add Manual Transaction</button>
            <div class="overflow-x-auto max-h-96 overflow-y-auto">
                <table class="w-full text-sm">
                    <thead class="sticky top-0 bg-gray-900 border-b border-cyan-400">
                        <tr><th>Receipt</th><th>Date & Time</th><th>Customer</th><th>Vehicle</th><th>Total</th><th>Payment</th><th>Actions</th></tr>
                    </thead>
                    <tbody>
                        ${appData.transactions.map(t => `
                            <tr class="border-b border-gray-700 hover:bg-gray-800 transition">
                                <td class="py-2 font-mono text-xs">${t.id}</td>
                                <td><span class="editable-field cursor-pointer" onclick="editTransactionDate('${t.id}')">${new Date(t.date).toLocaleString()}</span></td>
                                <td><span class="editable-field cursor-pointer" onclick="editTransactionCustomer('${t.id}')">${t.customer}</span></td>
                                <td>${t.vehicle} | ${t.plate}</td>
                                <td><span class="editable-field cursor-pointer text-green-400 font-bold" onclick="editTransactionAmount('${t.id}')">₱${t.total.toFixed(2)}</span></td>
                                <td>${t.paymentMethod || "Cash"}</td>
                                <td>
                                    <button onclick="editTransactionFull('${t.id}')" class="action-btn bg-cyan-600 text-white hover:bg-cyan-700"><i class="fas fa-edit"></i></button>
                                    <button onclick="deleteTransactionItem('${t.id}')" class="action-btn bg-red-600 text-white hover:bg-red-700"><i class="fas fa-trash"></i></button>
                                </td>
                            </tr>
                        `).join('')}
                    </tbody>
                    <tfoot class="sticky bottom-0 bg-gray-900 border-t border-cyan-400">
                        <tr><td colspan="4" class="py-2 text-right font-bold">GRAND TOTAL:</td><td class="py-2 text-green-400 font-bold">₱${totalIncome.toFixed(2)}</td><td colspan="2"></td></tr>
                    </tfoot>
                </table>
            </div>`;
    }

    window.showAddTransactionModal = () => {
        let modalHtml = `<div id="adminModal" class="modal"><div class="modal-content"><h3 class="text-xl font-bold mb-4">Add Manual Transaction</h3><div class="space-y-3"><div><label>Customer Name</label><input type="text" id="txCustomerInput" class="w-full bg-gray-800 p-2 rounded" value="Walk-in Customer"></div><div><label>Amount</label><input type="number" id="txAmountInput" class="w-full bg-gray-800 p-2 rounded"></div><div><label>Vehicle</label><input type="text" id="txVehicleInput" class="w-full bg-gray-800 p-2 rounded" placeholder="Toyota Vios"></div><div><label>Plate</label><input type="text" id="txPlateInput" class="w-full bg-gray-800 p-2 rounded"></div><div><label>Payment Method</label><select id="txPaymentInput" class="w-full bg-gray-800 p-2 rounded"><option>Cash</option><option>GCash</option><option>Maya</option><option>Credit Card</option><option>Bank Transfer</option></select></div><div class="flex gap-3 mt-4"><button onclick="confirmAddTransaction()" class="btn-neon flex-1 p-2 rounded">Add</button><button onclick="closeAdminModal()" class="bg-gray-700 flex-1 p-2 rounded">Cancel</button></div></div></div></div>`;
        document.body.insertAdjacentHTML('beforeend', modalHtml);
    };
    window.confirmAddTransaction = () => {
        const customer = document.getElementById('txCustomerInput')?.value || "Walk-in Customer";
        const amount = parseFloat(document.getElementById('txAmountInput')?.value);
        const vehicle = document.getElementById('txVehicleInput')?.value || "N/A";
        const plate = document.getElementById('txPlateInput')?.value || "N/A";
        const payment = document.getElementById('txPaymentInput')?.value || "Cash";
        if(isNaN(amount) || amount <= 0) { alert("Valid amount required!"); return; }
        const transaction = {
            id: generateReceiptNo(),
            date: new Date().toISOString(),
            customer: customer,
            vehicle: vehicle,
            plate: plate,
            items: [{ serviceName: "Manual Entry", variant: "Standard", price: amount, quantity: 1, total: amount, assignedStaff: [], remarks: "Manual transaction" }],
            subtotal: amount,
            discount: 0,
            loyaltyUsed: 0,
            total: amount,
            paymentMethod: payment,
            cashTendered: amount,
            change: 0,
            remarks: "Manual transaction added via admin"
        };
        appData.transactions.unshift(transaction);
        saveToLocal(); closeAdminModal(); renderApp();
    };
    window.editTransactionFull = (id) => {
        const t = appData.transactions.find(tx => tx.id === id);
        if(!t) return;
        let modalHtml = `<div id="adminModal" class="modal"><div class="modal-content"><h3 class="text-xl font-bold mb-4">Edit Transaction</h3><div class="space-y-3"><div><label>Customer</label><input type="text" id="editTxCustomer" class="w-full bg-gray-800 p-2 rounded" value="${t.customer}"></div><div><label>Amount</label><input type="number" id="editTxAmount" class="w-full bg-gray-800 p-2 rounded" value="${t.total}"></div><div><label>Vehicle</label><input type="text" id="editTxVehicle" class="w-full bg-gray-800 p-2 rounded" value="${t.vehicle}"></div><div><label>Plate</label><input type="text" id="editTxPlate" class="w-full bg-gray-800 p-2 rounded" value="${t.plate}"></div><div><label>Payment</label><select id="editTxPayment" class="w-full bg-gray-800 p-2 rounded"><option ${t.paymentMethod==="Cash"?"selected":""}>Cash</option><option ${t.paymentMethod==="GCash"?"selected":""}>GCash</option><option ${t.paymentMethod==="Maya"?"selected":""}>Maya</option><option ${t.paymentMethod==="Credit Card"?"selected":""}>Credit Card</option><option ${t.paymentMethod==="Bank Transfer"?"selected":""}>Bank Transfer</option></select></div><div class="flex gap-3 mt-4"><button onclick="confirmEditTransaction('${id}')" class="btn-neon flex-1 p-2 rounded">Update</button><button onclick="closeAdminModal()" class="bg-gray-700 flex-1 p-2 rounded">Cancel</button></div></div></div></div>`;
        document.body.insertAdjacentHTML('beforeend', modalHtml);
    };
    window.confirmEditTransaction = (id) => {
        const t = appData.transactions.find(tx => tx.id === id);
        if(!t) return;
        const customer = document.getElementById('editTxCustomer')?.value;
        const amount = parseFloat(document.getElementById('editTxAmount')?.value);
        const vehicle = document.getElementById('editTxVehicle')?.value;
        const plate = document.getElementById('editTxPlate')?.value;
        const payment = document.getElementById('editTxPayment')?.value;
        if(customer) t.customer = customer;
        if(!isNaN(amount) && amount > 0) { t.total = amount; t.subtotal = amount; }
        if(vehicle) t.vehicle = vehicle;
        if(plate) t.plate = plate;
        if(payment) t.paymentMethod = payment;
        saveToLocal(); closeAdminModal(); renderApp();
    };
    window.editTransactionAmount = (id) => { const t = appData.transactions.find(tx => tx.id === id); if(t) { const val = parseFloat(prompt("Edit amount:", t.total)); if(!isNaN(val) && val > 0) { t.total = val; t.subtotal = val; saveToLocal(); renderApp(); } } };
    window.editTransactionCustomer = (id) => { const t = appData.transactions.find(tx => tx.id === id); if(t) { const val = prompt("Edit customer:", t.customer); if(val !== null) { t.customer = val; saveToLocal(); renderApp(); } } };
    window.editTransactionDate = (id) => { const t = appData.transactions.find(tx => tx.id === id); if(t) { const val = prompt("Edit date (YYYY-MM-DD):", new Date(t.date).toISOString().slice(0,10)); if(val !== null) { t.date = new Date(val).toISOString(); saveToLocal(); renderApp(); } } };
    window.deleteTransactionItem = (id) => { if(confirm("Delete this transaction?")) { appData.transactions = appData.transactions.filter(t => t.id !== id); saveToLocal(); renderApp(); } };

    // ======================== ADMIN - COMMISSIONS ========================
    function renderAdminCommissions() {
        if(showLogin) return renderLogin();
        const commissionKeys = Object.keys(appData.commissionMatrix || {}).sort();
        const serviceNames = [...new Set(commissionKeys.map(k => k.split('_')[0]))];
        
        let tableHtml = '';
        serviceNames.forEach(serviceName => {
            const service = appData.services.find(s => s.name === serviceName);
            const sizes = service ? service.sizes : [];
            sizes.forEach(size => {
                const key = `${serviceName}_${size}`;
                const values = appData.commissionMatrix[key] || [0, 0, 0];
                tableHtml += `<tr class="border-b border-gray-700 hover:bg-gray-800 transition">
                    <td class="py-2 font-semibold">${serviceName}</td>
                    <td class="text-center">${size}</td>
                    <td class="text-center"><span class="editable-field cursor-pointer text-cyan-400" onclick="editCommissionValue('${key}', 0)">₱${values[0]}</span></td>
                    <td class="text-center"><span class="editable-field cursor-pointer text-green-400" onclick="editCommissionValue('${key}', 1)">₱${values[1]}</span></td>
                    <td class="text-center"><span class="editable-field cursor-pointer text-yellow-400" onclick="editCommissionValue('${key}', 2)">₱${values[2]}</span></td>
                    <td class="text-center"><button onclick="editCommissionFull('${key}')" class="action-btn bg-cyan-600 text-white hover:bg-cyan-700"><i class="fas fa-edit"></i></button>
                    <button onclick="deleteCommissionEntry('${key}')" class="action-btn bg-red-600 text-white hover:bg-red-700"><i class="fas fa-trash"></i></button></td>
                </tr>`;
            });
        });

        return `<h3 class="text-xl font-bold mb-4"><i class="fas fa-percent"></i> Commission Matrix Management</h3>
            <div class="flex gap-3 mb-4 flex-wrap">
                <button onclick="showAddCommissionModal()" class="bg-green-600 px-3 py-1 rounded"><i class="fas fa-plus"></i> Add Entry</button>
                <button onclick="autoGenerateCommissions()" class="bg-purple-600 px-3 py-1 rounded"><i class="fas fa-magic"></i> Auto-Generate</button>
                <button onclick="resetCommissionMatrix()" class="bg-red-600 px-3 py-1 rounded"><i class="fas fa-undo"></i> Reset</button>
            </div>
            <div class="overflow-x-auto max-h-96 overflow-y-auto">
                <table class="w-full text-sm">
                    <thead class="sticky top-0 bg-gray-900 border-b-2 border-cyan-400">
                        <tr><th class="text-left">Service</th><th class="text-center">Size</th><th class="text-center text-cyan-400">30%</th><th class="text-center text-green-400">35%</th><th class="text-center text-yellow-400">40%</th><th class="text-center">Actions</th></tr>
                    </thead>
                    <tbody>${tableHtml || '<tr><td colspan="6" class="text-center py-4 text-gray-400">No commission entries found.</td></tr>'}</tbody>
                    <tfoot class="sticky bottom-0 bg-gray-900 border-t border-cyan-400">
                        <tr><td colspan="6" class="py-2 text-center text-xs text-gray-400">Total Entries: ${commissionKeys.length}</td></tr>
                    </tfoot>
                </table>
            </div>`;
    }

    window.editCommissionValue = (key, tierIndex) => {
        const values = appData.commissionMatrix[key] || [0, 0, 0];
        const tierLabels = ['30%', '35%', '40%'];
        const newVal = parseFloat(prompt(`Edit ${tierLabels[tierIndex]} commission for ${key}:`, values[tierIndex]));
        if (!isNaN(newVal) && newVal >= 0) {
            values[tierIndex] = newVal;
            appData.commissionMatrix[key] = values;
            saveToLocal(); renderApp();
        }
    };
    window.editCommissionFull = (key) => {
        const values = appData.commissionMatrix[key] || [0, 0, 0];
        let modalHtml = `<div id="adminModal" class="modal"><div class="modal-content"><h3 class="text-xl font-bold mb-4">Edit Commission: ${key}</h3>
            <div class="space-y-3">
                <div><label class="text-cyan-400">30% Tier</label><input type="number" id="editComm30" class="w-full bg-gray-800 p-2 rounded" value="${values[0]}"></div>
                <div><label class="text-green-400">35% Tier</label><input type="number" id="editComm35" class="w-full bg-gray-800 p-2 rounded" value="${values[1]}"></div>
                <div><label class="text-yellow-400">40% Tier</label><input type="number" id="editComm40" class="w-full bg-gray-800 p-2 rounded" value="${values[2]}"></div>
                <div class="flex gap-3 mt-4"><button onclick="confirmEditCommission('${key}')" class="btn-neon flex-1 p-2 rounded">Update</button><button onclick="closeAdminModal()" class="bg-gray-700 flex-1 p-2 rounded">Cancel</button></div>
            </div></div></div>`;
        document.body.insertAdjacentHTML('beforeend', modalHtml);
    };
    window.confirmEditCommission = (key) => {
        const val30 = parseFloat(document.getElementById('editComm30')?.value);
        const val35 = parseFloat(document.getElementById('editComm35')?.value);
        const val40 = parseFloat(document.getElementById('editComm40')?.value);
        if (!isNaN(val30) && !isNaN(val35) && !isNaN(val40)) {
            appData.commissionMatrix[key] = [val30, val35, val40];
            saveToLocal(); closeAdminModal(); renderApp();
        }
    };
    window.deleteCommissionEntry = (key) => {
        if(confirm(`Delete commission entry for "${key}"?`)) {
            delete appData.commissionMatrix[key];
            saveToLocal(); renderApp();
        }
    };
    window.showAddCommissionModal = () => {
        const serviceNames = appData.services.map(s => s.name);
        const sizes = ['Small', 'Medium', 'Large', 'XL', '2XL', '3XL', 'Standard'];
        let modalHtml = `<div id="adminModal" class="modal"><div class="modal-content"><h3 class="text-xl font-bold mb-4">Add Commission Entry</h3>
            <div class="space-y-3">
                <div><label class="text-gray-400">Service Name</label><select id="addCommService" class="w-full bg-gray-800 p-2 rounded">${serviceNames.map(n => `<option>${n}</option>`).join('')}</select></div>
                <div><label class="text-gray-400">Size</label><select id="addCommSize" class="w-full bg-gray-800 p-2 rounded">${sizes.map(s => `<option>${s}</option>`).join('')}</select></div>
                <div><label class="text-cyan-400">30% Commission</label><input type="number" id="addComm30" class="w-full bg-gray-800 p-2 rounded"></div>
                <div><label class="text-green-400">35% Commission</label><input type="number" id="addComm35" class="w-full bg-gray-800 p-2 rounded"></div>
                <div><label class="text-yellow-400">40% Commission</label><input type="number" id="addComm40" class="w-full bg-gray-800 p-2 rounded"></div>
                <div class="flex gap-3 mt-4"><button onclick="confirmAddCommission()" class="btn-neon flex-1 p-2 rounded">Add</button><button onclick="closeAdminModal()" class="bg-gray-700 flex-1 p-2 rounded">Cancel</button></div>
            </div></div></div>`;
        document.body.insertAdjacentHTML('beforeend', modalHtml);
    };
    window.confirmAddCommission = () => {
        const service = document.getElementById('addCommService')?.value;
        const size = document.getElementById('addCommSize')?.value;
        const val30 = parseFloat(document.getElementById('addComm30')?.value);
        const val35 = parseFloat(document.getElementById('addComm35')?.value);
        const val40 = parseFloat(document.getElementById('addComm40')?.value);
        if (!service || !size) { alert("Service and size required!"); return; }
        const key = `${service}_${size}`;
        if (!isNaN(val30) && !isNaN(val35) && !isNaN(val40)) {
            appData.commissionMatrix[key] = [val30, val35, val40];
            saveToLocal(); closeAdminModal(); renderApp();
        }
    };
    window.autoGenerateCommissions = () => {
        if(confirm("Auto-generate commissions based on service prices?")) {
            appData.services.forEach(service => {
                service.sizes.forEach(size => {
                    const price = service.prices[size];
                    if (price) {
                        const tier30 = Math.round(price * 0.26);
                        const tier35 = Math.round(price * 0.30);
                        const tier40 = Math.round(price * 0.35);
                        const key = `${service.name}_${size}`;
                        appData.commissionMatrix[key] = [tier30, tier35, tier40];
                    }
                });
            });
            saveToLocal(); renderApp(); alert("Commissions auto-generated!");
        }
    };
    window.resetCommissionMatrix = () => {
        if(confirm("Reset commission matrix to default?")) {
            appData.commissionMatrix = initializeCommissionMatrix();
            saveToLocal(); renderApp();
        }
    };

    // ======================== ADMIN - USERS ========================
    function renderAdminUsers() {
        if(showLogin) return renderLogin();
        if(!hasRole(['Owner'])) {
            return `<div class="card-premium p-8 text-center"><i class="fas fa-lock text-4xl text-red-400 mb-4"></i><p class="text-xl">Access Denied. Owner role required.</p></div>`;
        }
        return `<h3 class="text-xl font-bold mb-4"><i class="fas fa-user-shield"></i> User Management</h3>
            <button onclick="addUser()" class="bg-green-600 px-3 py-1 rounded mb-4">+ Add User</button>
            <div class="overflow-x-auto"><table class="w-full text-sm"><thead class="border-b border-cyan-400"><tr><th>Username</th><th>Name</th><th>Role</th><th>Status</th><th>Actions</th></tr></thead><tbody>
            ${appData.users.map(u => `<tr class="border-b border-gray-700"><td class="py-2">${u.username}</td><td>${u.name}</td><td><span class="role-badge role-${u.role.toLowerCase()}">${u.role}</span></td><td><button onclick="toggleUserStatus(${u.id})" class="text-${u.isActive ? 'green' : 'red'}-400">${u.isActive ? 'Active' : 'Inactive'}</button></td><td><button onclick="deleteUser(${u.id})" class="text-red-400"><i class="fas fa-trash"></i></button></td></tr>`).join('')}</tbody></table></div>`;
    }
    window.addUser = () => {
        let username = prompt("Username:"); if(!username) return;
        let password = prompt("Password:") || "password123";
        let name = prompt("Full name:") || username;
        let role = prompt("Role (Owner/Manager/Cashier/Staff):") || "Staff";
        appData.users.push({ id: Date.now(), username, password, role, name, isActive: true });
        saveToLocal(); renderApp();
    };
    window.toggleUserStatus = (id) => { let u = appData.users.find(u=>u.id===id); if(u) { u.isActive = !u.isActive; saveToLocal(); renderApp(); } };
    window.deleteUser = (id) => { if(confirm("Delete user?")) { appData.users = appData.users.filter(u=>u.id!==id); saveToLocal(); renderApp(); } };

    // ======================== ADMIN - SYSTEM ========================
    function renderAdminSystem() {
        if(showLogin) return renderLogin();
        if(!hasRole(['Owner'])) {
            return `<div class="card-premium p-8 text-center"><i class="fas fa-lock text-4xl text-red-400 mb-4"></i><p class="text-xl">Access Denied. Owner role required.</p></div>`;
        }
        const s = appData.systemSettings;
        const logoHtml = s.companyLogo ? `<img src="${s.companyLogo}" class="company-logo" style="max-width:120px;max-height:80px;">` : 'No logo uploaded';
        const firebaseStatus = firebaseInitialized ? '✅ Firebase Connected' : '❌ Firebase Not Configured';
        
        return `<h3 class="text-xl font-bold mb-4"><i class="fas fa-cogs"></i> System Settings</h3>
            <div class="grid grid-cols-1 md:grid-cols-2 gap-4">
                <div class="space-y-3">
                    <div><label class="text-gray-400 text-sm">Company Logo</label>
                        <div class="flex items-center gap-3 bg-gray-800 p-3 rounded">
                            <div class="border border-gray-600 p-2 rounded">${logoHtml}</div>
                            <button onclick="uploadLogo()" class="bg-cyan-600 px-3 py-1 rounded text-sm"><i class="fas fa-upload"></i> Upload</button>
                            <button onclick="removeLogo()" class="bg-red-600 px-3 py-1 rounded text-sm"><i class="fas fa-times"></i> Remove</button>
                        </div>
                    </div>
                    <div><label class="text-gray-400 text-sm">Company Name</label><div class="editable-field bg-gray-800 p-2 rounded cursor-pointer" onclick="editSetting('companyName')">${s.companyName}</div></div>
                    <div><label class="text-gray-400 text-sm">Tagline</label><div class="editable-field bg-gray-800 p-2 rounded cursor-pointer" onclick="editSetting('tagline')">${s.tagline}</div></div>
                    <div><label class="text-gray-400 text-sm">Address</label><div class="editable-field bg-gray-800 p-2 rounded cursor-pointer" onclick="editSetting('address')">${s.address}</div></div>
                    <div><label class="text-gray-400 text-sm">Phone</label><div class="editable-field bg-gray-800 p-2 rounded cursor-pointer" onclick="editSetting('phone')">${s.phone}</div></div>
                    <div><label class="text-gray-400 text-sm">Facebook</label><div class="editable-field bg-gray-800 p-2 rounded cursor-pointer" onclick="editSetting('facebook')">${s.facebook}</div></div>
                    <div><label class="text-gray-400 text-sm">Payment Methods</label><div class="editable-field bg-gray-800 p-2 rounded cursor-pointer" onclick="editPaymentMethods()">${s.defaultPaymentMethods.join(', ')}</div></div>
                </div>
                <div class="space-y-3">
                    <div><label class="text-gray-400 text-sm">Receipt Footer</label><div class="editable-field bg-gray-800 p-2 rounded cursor-pointer" onclick="editSetting('receiptFooter')">${s.receiptFooter}</div></div>
                    <div><label class="text-gray-400 text-sm">Commission Split Rule</label><div class="editable-field bg-gray-800 p-2 rounded cursor-pointer" onclick="editSetting('commissionSplitRule')">${s.commissionSplitRule}</div></div>
                    <div><label class="text-gray-400 text-sm">Loyalty Points per ₱1000</label><div class="editable-field bg-gray-800 p-2 rounded cursor-pointer" onclick="editLoyalty()">${(s.loyaltyPointsPerPeso * 1000)} points</div></div>
                    <div><label class="text-gray-400 text-sm">Tax Rate</label><div class="editable-field bg-gray-800 p-2 rounded cursor-pointer" onclick="editTax()">${s.taxRate}%</div></div>
                    <div><label class="text-gray-400 text-sm">Business Hours</label><div class="editable-field bg-gray-800 p-2 rounded cursor-pointer" onclick="editSetting('businessHours')">${s.businessHours}</div></div>
                    <div><label class="text-gray-400 text-sm">Firebase Sync</label>
                        <div class="flex items-center gap-3 bg-gray-800 p-2 rounded">
                            <span class="text-sm">${firebaseStatus}</span>
                            <button onclick="toggleFirebaseSync()" class="bg-${s.firebaseSync ? 'red' : 'green'}-600 px-3 py-1 rounded text-sm text-white">
                                ${s.firebaseSync ? 'Disable Sync' : 'Enable Sync'}
                            </button>
                            ${s.firebaseSync && s.lastSync ? `<span class="text-xs text-gray-400">Last: ${new Date(s.lastSync).toLocaleString()}</span>` : ''}
                        </div>
                        ${!firebaseInitialized ? `<div class="text-xs text-yellow-400 mt-1">⚠️ Firebase not configured. Please add your Firebase config.</div>` : ''}
                    </div>
                </div>
            </div>
            <div class="mt-6 grid grid-cols-2 gap-3">
                <button onclick="backupData()" class="bg-cyan-600 p-2 rounded"><i class="fas fa-download"></i> Backup Database</button>
                <button onclick="restoreData()" class="bg-gray-600 p-2 rounded"><i class="fas fa-upload"></i> Restore Backup</button>
                <button onclick="exportAllData()" class="bg-green-600 p-2 rounded"><i class="fas fa-file-export"></i> Export All Data</button>
                <button onclick="resetSystem()" class="bg-red-600 p-2 rounded"><i class="fas fa-trash"></i> Reset to Default</button>
            </div>
            <div class="mt-4 p-3 bg-gray-800 rounded-lg text-xs">
                <div class="text-gray-400">System Info</div>
                <div>Version: 7.4 Enterprise (Firebase) | Last Saved: ${localStorage.getItem('8068_system_last_saved') ? new Date(localStorage.getItem('8068_system_last_saved')).toLocaleString() : 'Never'}</div>
                <div>Records: ${appData.transactions.length} TX | ${appData.customers.length} Customers | ${appData.staff.length} Staff</div>
                <div>Commission Entries: ${Object.keys(appData.commissionMatrix || {}).length}</div>
                <div>Open Services: ${appData.openServices.filter(s=>s.status!=='completed').length}</div>
                <div>Inventory Items: ${appData.inventory.length}</div>
                <div>Firebase Sync: ${s.firebaseSync ? '✅ Active' : '❌ Disabled'}</div>
                <div>Firebase Status: ${firebaseInitialized ? '✅ Connected' : '❌ Not Configured'}</div>
            </div>
        </div>`;
    }

    window.toggleFirebaseSync = () => {
        if (!firebaseInitialized) {
            alert('⚠️ Firebase is not configured. Please add your Firebase config in the code.');
            return;
        }
        
        appData.systemSettings.firebaseSync = !appData.systemSettings.firebaseSync;
        if (appData.systemSettings.firebaseSync) {
            const result = syncAllToFirebase();
            if (result) {
                alert('✅ Firebase Sync enabled! Data will be synced automatically to Firebase.');
            } else {
                alert('⚠️ Firebase Sync enabled but initial sync failed. Please check your connection.');
            }
        } else {
            alert('❌ Firebase Sync disabled.');
        }
        saveToLocal();
        renderApp();
    };

    window.editPaymentMethods = () => {
        const current = appData.systemSettings.defaultPaymentMethods.join(', ');
        const newMethods = prompt("Edit payment methods (comma separated):", current);
        if(newMethods) {
            appData.systemSettings.defaultPaymentMethods = newMethods.split(',').map(m => m.trim()).filter(m => m);
            saveToLocal(); renderApp();
        }
    };

    window.uploadLogo = () => {
        const input = document.createElement('input');
        input.type = 'file';
        input.accept = 'image/*';
        input.onchange = function(e) {
            const file = e.target.files[0];
            if (file) {
                const reader = new FileReader();
                reader.onload = function(event) {
                    appData.systemSettings.companyLogo = event.target.result;
                    saveToLocal();
                    renderApp();
                    alert('Logo uploaded successfully!');
                };
                reader.readAsDataURL(file);
            }
        };
        input.click();
    };
    window.removeLogo = () => {
        if(confirm('Remove company logo?')) {
            appData.systemSettings.companyLogo = null;
            saveToLocal();
            renderApp();
            alert('Logo removed.');
        }
    };
    window.editSetting = (key) => {
        const val = prompt(`Edit ${key}:`, appData.systemSettings[key]);
        if(val !== null) { appData.systemSettings[key] = val; saveToLocal(); renderApp(); }
    };
    window.editLoyalty = () => {
        const val = parseInt(prompt("Loyalty points per ₱1000 spent:", appData.systemSettings.loyaltyPointsPerPeso * 1000));
        if(!isNaN(val)) { appData.systemSettings.loyaltyPointsPerPeso = val / 1000; saveToLocal(); renderApp(); }
    };
    window.editTax = () => {
        const val = parseFloat(prompt("Tax rate (%):", appData.systemSettings.taxRate));
        if(!isNaN(val)) { appData.systemSettings.taxRate = val; saveToLocal(); renderApp(); }
    };
    window.backupData = () => {
        const dataStr = JSON.stringify(appData);
        localStorage.setItem("system_backup_8068", dataStr);
        const blob = new Blob([dataStr], {type: "application/json"});
        const a = document.createElement("a"); a.href = URL.createObjectURL(blob); a.download = `8068_backup_${new Date().toISOString().slice(0,10)}.json`; a.click();
        alert("Backup created!");
    };
    window.restoreData = () => {
        const backup = localStorage.getItem("system_backup_8068");
        if(backup) { try { Object.assign(appData, JSON.parse(backup)); saveToLocal(); alert("Restored!"); renderApp(); } catch(e) { alert("Backup corrupted!"); } }
        else { alert("No backup found!"); }
    };
    window.exportAllData = () => {
        const blob = new Blob([JSON.stringify(appData, null, 2)], {type: "application/json"});
        const a = document.createElement("a"); a.href = URL.createObjectURL(blob); a.download = `8068_full_export_${new Date().toISOString().slice(0,10)}.json`; a.click();
    };
    window.resetSystem = () => {
        if(confirm("WARNING: This will erase ALL data and reset to factory defaults. Continue?")) {
            if(confirm("Are you absolutely sure? This cannot be undone!")) { location.reload(); }
        }
    };

    window.closeAdminModal = () => document.getElementById('adminModal')?.remove();
    window.navigateAdmin = (panel) => { currentAdminPanel = panel; renderApp(); };

    // ======================== MAIN RENDER ========================
    function renderApp() {
        if(showLogin) {
            document.getElementById("mainContent").innerHTML = renderLogin();
            return;
        }
        
        const main = document.getElementById("mainContent");
        if(!main) return;
        let content = "";
        if(currentView === "dashboard") content = renderDashboard();
        else if(currentView === "pos") content = renderPOS();
        else if(currentView === "openservices") content = renderOpenServices();
        else if(currentView === "customers") content = renderCustomers();
        else if(currentView === "staff") content = renderStaffDashboard();
        else if(currentView === "reports") content = renderReports();
        else if(currentView === "admin") content = renderAdminDashboard();
        main.innerHTML = content;
        setTimeout(() => {
            document.getElementById('plateInput')?.addEventListener('input', (e) => vehicleInfo.plate = e.target.value);
            document.getElementById('modelInput')?.addEventListener('input', (e) => vehicleInfo.model = e.target.value);
        }, 50);
        updateClock();
        saveToLocal();
    }

    window.navigate = (view) => { currentView = view; renderApp(); };

    // ======================== LIVE CLOCK ========================
    function updateClock() {
        const now = new Date();
        const options = { year: 'numeric', month: 'short', day: 'numeric', hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: true };
        const clockEl = document.getElementById('liveClock');
        if(clockEl) clockEl.textContent = now.toLocaleDateString('en-US', options);
    }
    setInterval(updateClock, 1000);

    // ======================== INITIALIZE FIREBASE ========================
    const firebaseConfigured = initFirebase();
    if (firebaseConfigured) {
        updateFirebaseStatus('connecting');
        // Try to load data from Firebase on startup
        setTimeout(async () => {
            const loaded = await loadAllFromFirebase();
            if (loaded) {
                console.log('Data loaded from Firebase successfully!');
                renderApp();
            }
        }, 1000);
    } else {
        updateFirebaseStatus('disconnected');
    }

    // ======================== UI ========================
    document.body.innerHTML = `
        <div class="max-w-screen-2xl mx-auto px-4 py-5">
            <div class="flex flex-wrap justify-between items-center border-b border-blue-800 pb-4 mb-6">
                <div class="flex items-center gap-3">
                    ${appData.systemSettings.companyLogo ? `<img src="${appData.systemSettings.companyLogo}" class="company-logo" alt="Logo">` : ''}
                    <div><h1 class="text-3xl font-extrabold glow-text">${appData.systemSettings.companyName}</h1><p class="text-sm text-gray-400">${appData.systemSettings.tagline} | Enterprise v7.4</p></div>
                </div>
                <div class="flex gap-2 flex-wrap mt-2">
                    <button onclick="navigate('dashboard')" class="px-3 py-2 bg-gray-900 rounded-lg neon-border text-sm"><i class="fas fa-chart-line"></i> Dashboard</button>
                    <button onclick="navigate('pos')" class="px-3 py-2 bg-gray-900 rounded-lg neon-border text-sm"><i class="fas fa-cash-register"></i> POS</button>
                    <button onclick="navigate('openservices')" class="px-3 py-2 bg-orange-900 rounded-lg neon-border text-sm"><i class="fas fa-clipboard-list"></i> Open Services</button>
                    <button onclick="navigate('staff')" class="px-3 py-2 bg-gray-900 rounded-lg neon-border text-sm"><i class="fas fa-id-badge"></i> Staff</button>
                    <button onclick="navigate('reports')" class="px-3 py-2 bg-green-900 rounded-lg neon-border text-sm"><i class="fas fa-file-alt"></i> Reports</button>
                    <button onclick="navigate('customers')" class="px-3 py-2 bg-gray-900 rounded-lg neon-border text-sm"><i class="fas fa-users"></i> Customers</button>
                    <button onclick="navigate('admin')" class="px-3 py-2 bg-purple-900 rounded-lg neon-border text-sm"><i class="fas fa-crown"></i> Admin</button>
                </div>
            </div>
            <div id="mainContent"></div>
            <footer class="text-center text-gray-500 text-sm mt-12 pt-4 border-t border-gray-800">${appData.systemSettings.companyName} | ${appData.systemSettings.receiptFooter} | ${appData.systemSettings.businessHours}</footer>
        </div>
    `;
    renderApp();
    saveToLocal();

    console.log('%c8068 CAR SQUARE - Enterprise Management System v7.4 (Firebase)', 'color: #00d4ff; font-size: 20px; font-weight: bold;');
    console.log(`%cRecords: ${appData.transactions.length} TX | ${appData.customers.length} Customers | ${appData.staff.length} Staff | ${appData.openServices.filter(s=>s.status!=='completed').length} Open Services`, 'color: #00ff88;');
    console.log(`%cFirebase: ${firebaseInitialized ? '✅ Connected' : '❌ Not Configured'}`, 'color: ${firebaseInitialized ? '#00ff88' : '#ff6b6b'};');
</script>
</body>
</html>