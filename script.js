// Data dari Excel dan hasil scan
let inventoryData = [];
let scannedProducts = [];

// Variabel untuk scan detection
let scanTimeout = null;
let lastScanTime = 0;
const SCAN_DELAY = 500;

// Variabel untuk QR scanning
let qrStream = null;
let qrScannerActive = false;
let stopScanning = null;
let currentScannerType = 'barcode'; // 'barcode' atau 'qr'

// Elemen DOM
const uploadArea = document.getElementById('uploadArea');
const fileInput = document.getElementById('fileInput');
const fileName = document.getElementById('fileName');
const filePreview = document.getElementById('filePreview');
const previewTable = document.getElementById('previewTable');
const processBtn = document.getElementById('processBtn');
const scanSection = document.getElementById('scanSection');
const scanInput = document.getElementById('scanInput');
const scanBtn = document.getElementById('scanBtn');
const productList = document.getElementById('productList');
const compareBtn = document.getElementById('compareBtn');
const comparisonSection = document.getElementById('comparisonSection');
const comparisonBody = document.getElementById('comparisonBody');
const exportBtn = document.getElementById('exportBtn');
const resetBtn = document.getElementById('resetBtn');
const progressFill = document.getElementById('progressFill');
const statusText = document.getElementById('statusText');
const totalProducts = document.getElementById('totalProducts');
const scannedCount = document.getElementById('scannedCount');
const barcodeCount = document.getElementById('barcodeCount');
const totalDifference = document.getElementById('totalDifference');
const overstockCount = document.getElementById('overstockCount');
const understockCount = document.getElementById('understockCount');

// Elemen DOM baru untuk scan barcode
const scanIconBtn = document.getElementById('scanIconBtn');
const scanModal = document.getElementById('scanModal');
const closeModal = document.getElementById('closeModal');
const startScan = document.getElementById('startScan');
const stopScan = document.getElementById('stopScan');
const qrVideo = document.getElementById('qrVideo');
const qrCanvas = document.getElementById('qrCanvas');
const scanResult = document.getElementById('scanResult');
const scanArea = document.getElementById('scanArea');
const barcodeScanner = document.getElementById('barcodeScanner');
const scannerTypeRadios = document.querySelectorAll('input[name="scannerType"]');

// Event Listeners
fileInput.addEventListener('change', handleFileSelect);
processBtn.addEventListener('click', processExcelData);
scanBtn.addEventListener('click', addScannedProduct);
compareBtn.addEventListener('click', showComparison);
exportBtn.addEventListener('click', exportReport);
resetBtn.addEventListener('click', resetApp);

scanInput.addEventListener('input', handleScanInput);
scanInput.addEventListener('keydown', handleScanKeydown);

scanIconBtn.addEventListener('click', openScanModal);
closeModal.addEventListener('click', closeScanModal);
startScan.addEventListener('click', startQRScan);
stopScan.addEventListener('click', stopQRScan);

// Scanner type change
scannerTypeRadios.forEach(radio => {
    radio.addEventListener('change', function() {
        currentScannerType = this.value;
        console.log('Scanner type changed to:', currentScannerType);
        
        // Restart scan dengan type baru
        if (qrScannerActive) {
            stopQRScan();
            setTimeout(() => startQRScan(), 500);
        }
    });
});

// Drag and drop functionality
uploadArea.addEventListener('dragover', (e) => {
    e.preventDefault();
    uploadArea.classList.add('active');
});

uploadArea.addEventListener('dragleave', () => {
    uploadArea.classList.remove('active');
});

uploadArea.addEventListener('drop', (e) => {
    e.preventDefault();
    uploadArea.classList.remove('active');
    
    if (e.dataTransfer.files.length) {
        fileInput.files = e.dataTransfer.files;
        handleFileSelect();
    }
});

// ==================== FUNGSI SCAN BARCODE & QR ====================

function openScanModal() {
    scanModal.classList.remove('hidden');
    scanResult.textContent = 'Pilih jenis scanner dan klik "Mulai Scan"';
    scanResult.style.color = '#666';
    scanArea.classList.remove('scanning');
    
    // Reset scanner type
    currentScannerType = 'barcode';
    document.querySelector('input[name="scannerType"][value="barcode"]').checked = true;
}

function closeScanModal() {
    stopQRScan();
    scanModal.classList.add('hidden');
}

function startQRScan() {
    if (qrScannerActive) return;
    
    console.log('Starting scanner type:', currentScannerType);
    scanResult.textContent = 'Mengaktifkan kamera...';
    scanResult.style.color = '#666';
    scanArea.classList.add('scanning');
    
    // Reset semua element
    qrVideo.srcObject = null;
    qrVideo.classList.add('hidden');
    barcodeScanner.classList.add('hidden');
    
    if (!navigator.mediaDevices || !navigator.mediaDevices.getUserMedia) {
        scanResult.textContent = 'Browser tidak mendukung akses kamera';
        scanResult.style.color = 'var(--danger)';
        scanArea.classList.remove('scanning');
        return;
    }
    
    const constraints = {
        video: {
            facingMode: 'environment',
            width: { min: 640, ideal: 1280 },
            height: { min: 480, ideal: 720 }
        }
    };
    
    navigator.mediaDevices.getUserMedia(constraints)
    .then(function(stream) {
        console.log('Camera access granted for:', currentScannerType);
        qrStream = stream;
        qrScannerActive = true;
        
        if (currentScannerType === 'qr') {
            startQRCodeScanner(stream);
        } else {
            startBarcodeScanner(stream);
        }
        
    }).catch(function(error) {
        console.error('Camera access error:', error);
        scanArea.classList.remove('scanning');
        showCameraError(error);
    });
}

function startQRCodeScanner(stream) {
    qrVideo.srcObject = stream;
    qrVideo.classList.remove('hidden');
    
    qrVideo.onloadeddata = function() {
        console.log('QR Video data loaded');
        scanArea.classList.remove('scanning');
        scanResult.textContent = 'Arahkan kamera ke QR Code...';
        scanResult.style.color = 'var(--warning)';
        startScan.disabled = true;
        stopScan.disabled = false;
        
        qrVideo.play().then(() => {
            console.log('QR Video playing');
            stopScanning = scanQRCode();
        }).catch(playError => {
            console.error('QR Video play error:', playError);
            scanResult.textContent = 'Error memutar video: ' + playError.message;
            scanResult.style.color = 'var(--danger)';
        });
    };
}

function startBarcodeScanner(stream) {
    // Setup QuaggaJS untuk barcode
    barcodeScanner.classList.remove('hidden');
    
    Quagga.init({
        inputStream: {
            name: "Live",
            type: "LiveStream",
            target: barcodeScanner,
            constraints: {
                width: 640,
                height: 480,
                facingMode: "environment"
            }
        },
        decoder: {
            readers: [
                "code_128_reader",
                "ean_reader",
                "ean_8_reader",
                "code_39_reader",
                "code_39_vin_reader",
                "codabar_reader",
                "upc_reader",
                "upc_e_reader"
            ]
        },
        locator: {
            patchSize: "medium",
            halfSample: true
        },
        locate: true,
        numOfWorkers: 2
    }, function(err) {
        if (err) {
            console.error('Quagga init error:', err);
            scanResult.textContent = 'Error inisialisasi barcode scanner: ' + err.message;
            scanResult.style.color = 'var(--danger)';
            scanArea.classList.remove('scanning');
            return;
        }
        
        console.log('Quagga initialized successfully');
        scanArea.classList.remove('scanning');
        scanResult.textContent = 'Arahkan kamera ke Barcode...';
        scanResult.style.color = 'var(--warning)';
        startScan.disabled = true;
        stopScan.disabled = false;
        
        Quagga.start();
        
        // Listen for barcode detection
        Quagga.onDetected(function(result) {
            if (result && result.codeResult) {
                console.log('Barcode detected:', result.codeResult.code);
                handleScanResult(result.codeResult.code);
            }
        });
    });
}

function stopQRScan() {
    console.log('Stopping scanner...');
    qrScannerActive = false;
    
    // Stop scanning loop untuk QR
    if (stopScanning) {
        stopScanning();
        stopScanning = null;
    }
    
    // Stop Quagga untuk barcode
    if (currentScannerType === 'barcode' && typeof Quagga !== 'undefined') {
        Quagga.stop();
    }
    
    // Stop camera stream
    if (qrStream) {
        qrStream.getTracks().forEach(track => {
            track.stop();
        });
        qrStream = null;
    }
    
    // Reset video
    qrVideo.srcObject = null;
    
    startScan.disabled = false;
    stopScan.disabled = true;
    scanResult.textContent = 'Scan dihentikan';
    scanResult.style.color = '#666';
    scanArea.classList.remove('scanning');
}

function scanQRCode() {
    if (!qrScannerActive) return;
    
    console.log('Starting QR scanning loop...');
    const canvasContext = qrCanvas.getContext('2d');
    let scanning = true;
    
    function scanFrame() {
        if (!scanning || !qrScannerActive) return;
        
        try {
            if (qrVideo.readyState >= qrVideo.HAVE_CURRENT_DATA && qrVideo.videoWidth > 0) {
                qrCanvas.width = qrVideo.videoWidth;
                qrCanvas.height = qrVideo.videoHeight;
                
                canvasContext.drawImage(qrVideo, 0, 0, qrCanvas.width, qrCanvas.height);
                
                const imageData = canvasContext.getImageData(0, 0, qrCanvas.width, qrCanvas.height);
                
                if (typeof jsQR !== 'undefined') {
                    const code = jsQR(imageData.data, imageData.width, imageData.height, {
                        inversionAttempts: 'dontInvert',
                    });
                    
                    if (code) {
                        console.log('QR Code found:', code.data);
                        handleScanResult(code.data);
                        return;
                    }
                }
            }
            
            if (scanning) {
                requestAnimationFrame(scanFrame);
            }
        } catch (error) {
            console.error('Scan frame error:', error);
            if (scanning) {
                requestAnimationFrame(scanFrame);
            }
        }
    }
    
    scanFrame();
    
    return () => {
        scanning = false;
        console.log('QR Scanning loop stopped');
    };
}

function handleScanResult(result) {
    console.log('Processing scan result:', result);
    
    scanResult.textContent = `Berhasil scan: ${result}`;
    scanResult.style.color = 'var(--success)';
    
    // Auto-input ke field scan
    scanInput.value = result;
    
    // Auto-process setelah delay
    setTimeout(() => {
        const product = findProductBySKUOrBarcode(result);
        if (product) {
            addProductToScanned(product);
            scanInput.value = '';
            
            // Tutup modal setelah berhasil
            setTimeout(() => {
                closeScanModal();
                scanResult.textContent = 'Scan berhasil! Produk telah ditambahkan.';
            }, 1000);
        } else {
            scanResult.textContent = `Produk dengan kode "${result}" tidak ditemukan`;
            scanResult.style.color = 'var(--danger)';
        }
    }, 500);
}

function showCameraError(error) {
    let errorMessage = 'Gagal mengakses kamera: ';
    
    if (error.name === 'NotAllowedError') {
        errorMessage += 'Izin kamera ditolak. Silakan izinkan akses kamera.';
    } else if (error.name === 'NotFoundError') {
        errorMessage += 'Tidak ada kamera yang ditemukan.';
    } else if (error.name === 'NotSupportedError') {
        errorMessage += 'Browser tidak mendukung fitur ini.';
    } else if (error.name === 'NotReadableError') {
        errorMessage += 'Kamera sedang digunakan aplikasi lain.';
    } else {
        errorMessage += error.message;
    }
    
    scanResult.textContent = errorMessage;
    scanResult.style.color = 'var(--danger)';
    console.error('Camera error:', error);
}

scanModal.addEventListener('click', function(e) {
    if (e.target === scanModal) {
        closeScanModal();
    }
});

// ==================== FUNGSI UTAMA (TETAP SAMA) ====================

function handleScanInput(e) {
    const value = e.target.value.trim();
    
    if (scanTimeout) {
        clearTimeout(scanTimeout);
    }
    
    const now = Date.now();
    const timeSinceLastKey = now - lastScanTime;
    lastScanTime = now;
    
    if (timeSinceLastKey < 100) {
        scanTimeout = setTimeout(() => {
            performAutoScan(value);
        }, SCAN_DELAY);
    } else {
        scanTimeout = setTimeout(() => {
            performAutoScan(value);
        }, 800);
    }
}

function handleScanKeydown(e) {
    if (e.key === 'Enter') {
        if (scanTimeout) {
            clearTimeout(scanTimeout);
        }
        addScannedProduct();
    }
}

function performAutoScan(value) {
    if (!value) return;
    
    const product = findProductBySKUOrBarcode(value);
    
    if (product) {
        addProductToScanned(product);
        scanInput.value = '';
    }
}

function findProductBySKUOrBarcode(searchValue) {
    const cleanSearch = cleanValue(searchValue);
    
    let product = inventoryData.find(item => 
        cleanValue(item.sku) === cleanSearch
    );
    
    if (!product) {
        product = inventoryData.find(item => 
            item.barcode && cleanValue(item.barcode) === cleanSearch
        );
    }
    
    if (!product) {
        product = inventoryData.find(item => 
            cleanValue(item.sku).includes(cleanSearch) || 
            cleanSearch.includes(cleanValue(item.sku))
        );
    }
    
    if (!product) {
        product = inventoryData.find(item => 
            item.barcode && (
                cleanValue(item.barcode).includes(cleanSearch) || 
                cleanSearch.includes(cleanValue(item.barcode))
            )
        );
    }
    
    return product;
}

// ==================== FUNGSI YANG DIUPDATE ====================

function addProductToScanned(product) {
    const existingIndex = scannedProducts.findIndex(item => item.sku === product.sku);
    
    if (existingIndex !== -1) {
        // Jika produk sudah ada, tingkatkan jumlah dan pindahkan ke atas
        scannedProducts[existingIndex].scannedCount += 1;
        
        // Pindahkan produk ke paling atas
        const existingProduct = scannedProducts[existingIndex];
        scannedProducts.splice(existingIndex, 1);
        scannedProducts.unshift(existingProduct);
    } else {
        // Jika produk baru, tambahkan di paling atas
        scannedProducts.unshift({
            id: scannedProducts.length + 1,
            title: product.title,
            store: product.store,
            sku: product.sku,
            barcode: product.barcode,
            systemStock: product.stock,
            scannedCount: 1
        });
    }
    
    updateProductList();
    scannedCount.textContent = scannedProducts.length;
    compareBtn.disabled = false;
    updateProgress(75, `${scannedProducts.length} produk berhasil di-scan`);
}

function updateProductList() {
    productList.innerHTML = '';
    
    if (scannedProducts.length === 0) {
        productList.innerHTML = '<p style="text-align: center; color: #666;">Belum ada barang yang di-scan</p>';
        return;
    }
    
    scannedProducts.forEach((product, index) => {
        const productItem = document.createElement('div');
        productItem.className = 'product-item';
        
        // Tambahkan highlight untuk produk terbaru (paling atas)
        if (index === 0) {
            productItem.style.borderLeft = '4px solid var(--success)';
            productItem.style.background = 'rgba(39, 174, 96, 0.05)';
        }
        
        productItem.innerHTML = `
            <div class="product-info">
                <div class="product-title">${product.title}</div>
                <div class="product-sku">SKU: ${product.sku}</div>
                ${product.barcode ? `<div class="product-barcode">Barcode: ${product.barcode}</div>` : ''}
                <div class="product-store">Toko: ${product.store}</div>
            </div>
            <div style="display: flex; align-items: center;">
                <span class="product-stock">Stok Sistem: ${product.systemStock}</span>
                <input type="number" class="product-input" value="${product.scannedCount}" min="0" 
                    onchange="updateScannedCount('${product.sku}', this.value)">
                <button class="btn btn-danger" onclick="removeProduct('${product.sku}')" style="margin-left: 10px;">Hapus</button>
            </div>
        `;
        
        productList.appendChild(productItem);
    });
}

function updateScannedCount(sku, value) {
    const newCount = parseInt(value);
    const index = scannedProducts.findIndex(item => item.sku === sku);
    
    if (index !== -1 && !isNaN(newCount) && newCount >= 0) {
        scannedProducts[index].scannedCount = newCount;
        
        // Pindahkan produk ke atas saat jumlah diupdate
        const product = scannedProducts[index];
        scannedProducts.splice(index, 1);
        scannedProducts.unshift(product);
        updateProductList();
    } else {
        // Reset ke nilai sebelumnya jika input tidak valid
        updateProductList();
    }
}

function removeProduct(sku) {
    const index = scannedProducts.findIndex(item => item.sku === sku);
    if (index !== -1) {
        scannedProducts.splice(index, 1);
        updateProductList();
        scannedCount.textContent = scannedProducts.length;
        
        if (scannedProducts.length === 0) {
            compareBtn.disabled = true;
        }
        
        updateProgress(66, `${scannedProducts.length} produk terscan`);
    }
}

// ==================== FUNGSI LAIN TETAP SAMA ====================

function handleFileSelect() {
    if (fileInput.files.length > 0) {
        const file = fileInput.files[0];
        fileName.textContent = `File dipilih: ${file.name}`;
        readExcelFile(file);
    }
}

function readExcelFile(file) {
    const reader = new FileReader();
    
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
            processExcelDataToTable(jsonData);
        } catch (error) {
            console.error('Error reading Excel file:', error);
            alert('Error membaca file Excel. Pastikan format file benar.');
        }
    };
    
    reader.onerror = function() {
        alert('Error membaca file');
    };
    
    reader.readAsArrayBuffer(file);
}

function cleanValue(value) {
    if (value === null || value === undefined) return '';
    let cleaned = value.toString().trim();
    if (!isNaN(cleaned) && cleaned.includes('.')) {
        cleaned = cleaned.replace(/\.0+$/, '');
    }
    return cleaned;
}

function processExcelDataToTable(data) {
    previewTable.querySelector('tbody').innerHTML = '';
    inventoryData = [];
    
    let headerRowIndex = -1;
    for (let i = 0; i < Math.min(10, data.length); i++) {
        const row = data[i];
        if (row && row.length >= 4) {
            const headers = row.map(h => h ? cleanValue(h).toLowerCase() : '');
            const hasProduct = headers.some(h => h.includes('judul') && h.includes('produk'));
            const hasStore = headers.some(h => h.includes('toko'));
            const hasSKU = headers.some(h => h.includes('sku'));
            const hasStock = headers.some(h => h.includes('unit') && h.includes('persediaan'));
            const hasBarcode = headers.some(h => h.includes('barcode'));
            
            if (hasProduct && hasStore && hasSKU && hasStock) {
                headerRowIndex = i;
                break;
            }
        }
    }
    
    if (headerRowIndex === -1) {
        for (let i = 0; i < Math.min(5, data.length); i++) {
            const row = data[i];
            if (row && row.length >= 4) {
                headerRowIndex = i;
                break;
            }
        }
    }
    
    if (headerRowIndex === -1) {
        alert('Tidak dapat menemukan header yang sesuai dalam file Excel. Pastikan file memiliki kolom: Judul Produk, Toko, Nomor SKU, Unit dalam Persediaan');
        return;
    }
    
    const headerRow = data[headerRowIndex].map(h => cleanValue(h).toLowerCase());
    const titleIndex = headerRow.findIndex(h => h.includes('judul') || h.includes('produk'));
    const storeIndex = headerRow.findIndex(h => h.includes('toko') || h.includes('store'));
    const skuIndex = headerRow.findIndex(h => h.includes('sku') || h.includes('nomor'));
    const barcodeIndex = headerRow.findIndex(h => h.includes('barcode'));
    const stockIndex = headerRow.findIndex(h => h.includes('unit') || h.includes('persediaan') || h.includes('stock'));
    
    if (skuIndex === -1) {
        alert('Kolom "Nomor SKU" tidak ditemukan dalam file Excel.');
        return;
    }
    
    const dataRows = data.slice(headerRowIndex + 1);
    let validDataCount = 0;
    let productsWithBarcode = 0;
    
    dataRows.forEach((row, index) => {
        if (row && row.length > Math.max(titleIndex, storeIndex, skuIndex, barcodeIndex, stockIndex)) {
            const skuValue = cleanValue(row[skuIndex]);
            const barcodeValue = barcodeIndex !== -1 ? cleanValue(row[barcodeIndex]) : '';
            
            if (skuValue || barcodeValue) {
                const product = {
                    title: titleIndex !== -1 ? cleanValue(row[titleIndex]) : `Produk ${validDataCount + 1}`,
                    store: storeIndex !== -1 ? cleanValue(row[storeIndex]) : 'Toko Tidak Diketahui',
                    sku: skuValue,
                    barcode: barcodeValue,
                    stock: stockIndex !== -1 ? parseInt(cleanValue(row[stockIndex])) || 0 : 0
                };
                
                const tableRow = document.createElement('tr');
                tableRow.innerHTML = `
                    <td>${product.title}</td>
                    <td>${product.store}</td>
                    <td>${product.sku}</td>
                    <td>${product.barcode || '-'}</td>
                    <td>${product.stock}</td>
                `;
                previewTable.querySelector('tbody').appendChild(tableRow);
                
                validDataCount++;
                if (product.barcode) productsWithBarcode++;
                inventoryData.push(product);
            }
        }
    });
    
    if (validDataCount > 0) {
        filePreview.classList.remove('hidden');
        processBtn.disabled = false;
        updateProgress(33, `File Excel berhasil dibaca. ${validDataCount} produk ditemukan.`);
        totalProducts.textContent = validDataCount;
        barcodeCount.textContent = productsWithBarcode;
    } else {
        alert('Tidak ada data yang valid ditemukan dalam file Excel. Pastikan file memiliki data dengan format yang benar.');
    }
}

function processExcelData() {
    if (inventoryData.length === 0) {
        alert('Tidak ada data untuk diproses');
        return;
    }
    
    updateProgress(66, 'Data Excel berhasil diproses. Siap untuk scan barang.');
    scanSection.classList.remove('hidden');
    scanInput.focus();
}

function addScannedProduct() {
    const sku = scanInput.value.trim();
    
    if (!sku) {
        alert('Masukkan kode SKU terlebih dahulu');
        return;
    }
    
    const product = findProductBySKUOrBarcode(sku);
    
    if (!product) {
        alert(`Produk dengan kode "${sku}" tidak ditemukan dalam data Excel. Pastikan kode sesuai dengan yang ada di file.`);
        scanInput.value = '';
        return;
    }
    
    addProductToScanned(product);
    scanInput.value = '';
    scanInput.focus();
}

function showComparison() {
    comparisonBody.innerHTML = '';
    
    let totalDiff = 0;
    let overstock = 0;
    let understock = 0;
    
    scannedProducts.forEach(product => {
        const difference = product.scannedCount - product.systemStock;
        totalDiff += difference;
        
        if (difference > 0) overstock++;
        if (difference < 0) understock++;
        
        const differenceClass = difference > 0 ? 'positive' : difference < 0 ? 'negative' : 'neutral';
        const differenceText = difference > 0 ? `+${difference}` : difference;
        
        let statusText, statusClass;
        if (difference === 0) {
            statusText = 'Sesuai';
            statusClass = 'status-match';
        } else if (difference > 0) {
            statusText = 'Lebih';
            statusClass = 'status-overstock';
        } else {
            statusText = 'Kurang';
            statusClass = 'status-understock';
        }
        
        const row = document.createElement('tr');
        row.innerHTML = `
            <td>${product.title}</td>
            <td>${product.sku}</td>
            <td>${product.systemStock}</td>
            <td>${product.scannedCount}</td>
            <td class="${differenceClass}">${differenceText}</td>
            <td><span class="status-badge ${statusClass}">${statusText}</span></td>
        `;
        
        comparisonBody.appendChild(row);
    });
    
    totalDifference.textContent = totalDiff > 0 ? `+${totalDiff}` : totalDiff;
    totalDifference.className = totalDiff > 0 ? 'summary-value positive' : 
                               totalDiff < 0 ? 'summary-value negative' : 'summary-value';
    
    overstockCount.textContent = overstock;
    understockCount.textContent = understock;
    
    comparisonSection.classList.remove('hidden');
    updateProgress(100, 'Perbandingan stok selesai');
}

function exportReport() {
    if (scannedProducts.length === 0) {
        alert('Tidak ada data untuk diekspor');
        return;
    }
    
    const wb = XLSX.utils.book_new();
    const excelData = [
        ['LAPORAN PERBANDINGAN STOK PASLAM'],
        ['Tanggal', new Date().toLocaleDateString('id-ID')],
        [],
        ['Judul Produk', 'Nomor SKU', 'Barcode', 'Toko', 'Stok Sistem', 'Stok Fisik', 'Selisih', 'Status']
    ];
    
    scannedProducts.forEach(product => {
        const difference = product.scannedCount - product.systemStock;
        const status = difference === 0 ? 'Sesuai' : difference > 0 ? 'Lebih' : 'Kurang';
        
        excelData.push([
            product.title,
            product.sku,
            product.barcode || '',
            product.store,
            product.systemStock,
            product.scannedCount,
            difference,
            status
        ]);
    });
    
    excelData.push([]);
    excelData.push(['SUMMARY', '', '', '', '', '', '', '']);
    const totalDiff = scannedProducts.reduce((sum, product) => sum + (product.scannedCount - product.systemStock), 0);
    const overstock = scannedProducts.filter(product => product.scannedCount > product.systemStock).length;
    const understock = scannedProducts.filter(product => product.scannedCount < product.systemStock).length;
    
    excelData.push(['Total Selisih', '', '', '', '', '', totalDiff, '']);
    excelData.push(['Barang Lebih', '', '', '', '', '', overstock, '']);
    excelData.push(['Barang Kurang', '', '', '', '', '', understock, '']);
    
    const ws = XLSX.utils.aoa_to_sheet(excelData);
    XLSX.utils.book_append_sheet(wb, ws, 'Laporan Stok');
    XLSX.writeFile(wb, `Laporan_Stok_Paslam_${new Date().toISOString().split('T')[0]}.xlsx`);
    
    alert('Laporan berhasil diekspor!');
}

function resetApp() {
    inventoryData = [];
    scannedProducts = [];
    
    if (scanTimeout) {
        clearTimeout(scanTimeout);
        scanTimeout = null;
    }
    lastScanTime = 0;
    
    stopQRScan();
    
    fileInput.value = '';
    fileName.textContent = 'Belum ada file yang dipilih';
    filePreview.classList.add('hidden');
    previewTable.querySelector('tbody').innerHTML = '';
    processBtn.disabled = true;
    scanSection.classList.add('hidden');
    scanInput.value = '';
    productList.innerHTML = '';
    compareBtn.disabled = true;
    comparisonSection.classList.add('hidden');
    comparisonBody.innerHTML = '';
    
    totalProducts.textContent = '0';
    scannedCount.textContent = '0';
    barcodeCount.textContent = '0';
    totalDifference.textContent = '0';
    overstockCount.textContent = '0';
    understockCount.textContent = '0';
    
    updateProgress(0, 'Siap untuk upload data Excel');
}

function updateProgress(percentage, message) {
    progressFill.style.width = `${percentage}%`;
    statusText.textContent = message;
    
    const statusIndicator = document.querySelector('.status-indicator');
    statusIndicator.className = 'status-indicator';
    
    if (percentage === 100) {
        statusIndicator.classList.add('status-completed');
    } else if (percentage > 0) {
        statusIndicator.classList.add('status-pending');
    }
}

document.addEventListener('DOMContentLoaded', function() {
    console.log('Aplikasi Inventori Paslam siap digunakan');
    updateProgress(0, 'Siap untuk upload data Excel');
});
