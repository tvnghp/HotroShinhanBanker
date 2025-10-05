let excelData = [];
let recognition;

// Biến toàn cục cho việc chọn vùng ảnh
let isSelecting = false;
let isMoving = false;
let isResizing = false;
let startX, startY;
let selectionOverlay = null;
let currentImage = null;
let currentHandle = null;
let originalX, originalY, originalWidth, originalHeight;

// Khởi tạo đối tượng nhận diện giọng nói
function initSpeechRecognition() {
    // Tạo đối tượng nhận diện giọng nói nếu hỗ trợ
    if ('SpeechRecognition' in window || 'webkitSpeechRecognition' in window) {
        recognition = new (window.SpeechRecognition || window.webkitSpeechRecognition)();
        recognition.lang = 'vi-VN';
        recognition.interimResults = false;
        recognition.maxAlternatives = 1;
        
        // Xử lý kết quả nhận diện
        recognition.onresult = (event) => {
            const transcript = event.results[0][0].transcript.toLowerCase();
            const cleanedTranscript = transcript.replace(/[.,!?;:]/g, '').trim();
            console.log('Văn bản đã làm sạch:', cleanedTranscript);
            console.log('Nhận diện văn bản:', transcript);
            
            // Kiểm tra nếu là lệnh xóa
            if (transcript.includes('xóa') || transcript.includes('xoá') || 
                transcript.includes('làm mới') || transcript.includes('xóa kết quả')) {
                clearResults();
                showNotification('Đã xóa kết quả theo lệnh giọng nói');
            } 
            // Kiểm tra nếu là lệnh hiển thị tất cả (phần trăm phần trăm)
            else if (transcript.includes('phần trăm phần trăm') || 
                     transcript.includes('hiển thị tất cả') || 
                     transcript.includes('hiện tất cả') ||
                     transcript.includes('tất cả') ||
                     transcript.includes('hiện hết') ||
                     transcript.includes('xem hết')) {
                document.getElementById('searchInput').value = '%%';
                searchExcelData('%%');
                showNotification('Hiển thị tất cả dữ liệu theo yêu cầu');
            } else {
                // Nếu không phải lệnh đặc biệt, thì xử lý như tìm kiếm bình thường
                document.getElementById('searchInput').value = cleanedTranscript;
                searchExcelData(cleanedTranscript);
            }
            
            // Khôi phục nút sau khi có kết quả
            resetListeningButton();
        };

        recognition.onerror = (event) => {
            console.error('Lỗi trong quá trình nhận diện:', event.error);
            resetListeningButton();
        };
        
        recognition.onend = () => {
            resetListeningButton();
        };
    } else {
        showNotification('Trình duyệt của bạn không hỗ trợ nhận diện giọng nói', 'error');
        document.getElementById('startListeningButton').disabled = true;
    }
}

// Hàm để khởi động nhận diện giọng nói
function startListening() {
    if (!recognition) {
        initSpeechRecognition();
    }
    
    try {
        recognition.start();
        // Thêm thông báo đang lắng nghe
        document.getElementById('startListeningButton').textContent = 'Đang lắng nghe...';
        document.getElementById('startListeningButton').style.backgroundColor = '#4a9b4a';
        
        // Hiển thị hướng dẫn lệnh giọng nói
        showVoiceCommands();
    } catch (error) {
        console.error('Lỗi khi bắt đầu nhận diện:', error);
        resetListeningButton();
    }
}

// Khôi phục trạng thái nút lắng nghe
function resetListeningButton() {
    document.getElementById('startListeningButton').textContent = 'Bắt Đầu Nói';
    document.getElementById('startListeningButton').style.backgroundColor = '#8A1538';
    // Ẩn hướng dẫn lệnh giọng nói sau khi hoàn thành
    hideVoiceCommands();
}

// Hiển thị hướng dẫn lệnh giọng nói
function showVoiceCommands() {
    let voiceCommandsHelp = document.getElementById('voiceCommandsHelp');
    
    if (!voiceCommandsHelp) {
        voiceCommandsHelp = document.createElement('div');
        voiceCommandsHelp.id = 'voiceCommandsHelp';
        voiceCommandsHelp.innerHTML = `
            <div style="font-size: 13px; color: #666; margin-top: 5px; padding: 5px 10px; background-color: #f0f2f5; border-radius: 15px; display: inline-block;">
                Nói <strong>"xóa"</strong> hoặc <strong>"xóa kết quả"</strong> để làm mới kết quả<br>
                Nói <strong>"tất cả"</strong> hoặc <strong>"hiện tất cả"</strong> để hiển thị toàn bộ dữ liệu
            </div>
        `;
        
        // Thêm sau nút lắng nghe
        const controlsContainer = document.getElementById('controls-container');
        controlsContainer.appendChild(voiceCommandsHelp);
    } else {
        voiceCommandsHelp.style.display = 'block';
    }
}

// Ẩn hướng dẫn lệnh giọng nói
function hideVoiceCommands() {
    const voiceCommandsHelp = document.getElementById('voiceCommandsHelp');
    if (voiceCommandsHelp) {
        voiceCommandsHelp.style.display = 'none';
    }
}

// Xóa kết quả tìm kiếm và làm trống ô tìm kiếm
function clearResults() {
    document.getElementById('searchInput').value = '';
    document.getElementById('resultArea').innerHTML = '';
}

// Hàm để tìm kiếm dữ liệu trong file Excel
function searchExcelData(searchTerm) {
    if (excelData.length === 0) {
        showNotification('Vui lòng chọn file Excel trước khi tìm kiếm.', 'warning');
        return;
    }
    
    if (!searchTerm || searchTerm.trim() === '') {
        showNotification('Vui lòng nhập từ khóa tìm kiếm.', 'info');
        return;
    }

    let results;
    const isShowAll = searchTerm.trim() === '%%';
    
    // Nếu searchTerm là "%%", hiển thị tất cả dữ liệu
    if (isShowAll) {
        results = excelData.filter(row => row.some(cell => cell !== null && cell !== undefined && cell.toString().trim() !== ''));
    } else {
        // Tìm kiếm thông thường
        results = excelData.filter(row => 
            row.some(cell => cell && cell.toString().toLowerCase().includes(searchTerm.toLowerCase()))
        );
    }

    const resultArea = document.getElementById('resultArea');
    resultArea.innerHTML = ''; // Xóa kết quả cũ

    if (results.length > 0) {
        // Hiển thị số lượng kết quả tìm thấy
        const resultCount = document.createElement('div');
        resultCount.className = 'result-count';
        
        if (isShowAll) {
            resultCount.textContent = `Hiển thị tất cả ${results.length} bản ghi dữ liệu`;
        } else {
            resultCount.textContent = `Tìm thấy ${results.length} kết quả cho "${searchTerm}"`;
        }
        
        resultCount.style.marginBottom = '15px';
        resultCount.style.fontWeight = 'bold';
        resultCount.style.color = '#8A1538';
        resultCount.style.borderBottom = '1px solid #eaeaea';
        resultCount.style.paddingBottom = '8px';
        resultArea.appendChild(resultCount);
        
        // Hiển thị các kết quả tìm kiếm
        results.forEach((row, index) => {
            const resultRow = document.createElement('div');
            
            // Format dữ liệu kết quả để dễ đọc hơn
            const formattedData = formatResultRow(row, isShowAll ? "" : searchTerm);
            resultRow.innerHTML = formattedData;
            
            // Thêm số thứ tự cho mỗi kết quả
            resultRow.dataset.index = index + 1;
            
            resultArea.appendChild(resultRow);
        });
    } else {
        const noResult = document.createElement('div');
        noResult.className = 'no-result';
        noResult.innerHTML = `<p style="text-align: center; color: #888;">Không tìm thấy kết quả cho "${searchTerm}"</p>`;
        resultArea.appendChild(noResult);
    }
}

// Hàm định dạng kết quả để dễ đọc hơn và chỉ hiển thị cột F
function formatResultRow(row, searchTerm) {
    let result = '';
    
    // Chỉ lấy và hiển thị giá trị từ cột F (index 5)
    const cellF = row[5];
    
    if (cellF !== null && cellF !== undefined && cellF.toString().trim() !== '') {
        const cellStr = cellF.toString();
        const lowerCellStr = cellStr.toLowerCase();
        const lowerSearchTerm = searchTerm.toLowerCase();
        
        let formattedCell = cellStr;
        
        // Highlight từ khóa tìm kiếm nếu có
        if (lowerSearchTerm && lowerSearchTerm !== '%%' && lowerCellStr.includes(lowerSearchTerm)) {
            const startIndex = lowerCellStr.indexOf(lowerSearchTerm);
            const endIndex = startIndex + searchTerm.length;
            
            formattedCell = cellStr.substring(0, startIndex) + 
                   `<span style="background-color: #ffe2e8; font-weight: bold;">${cellStr.substring(startIndex, endIndex)}</span>` + 
                   cellStr.substring(endIndex);
        }
        
        result = `<span class="column-f-highlight" style="color: white; font-weight: bold; font-size: 1.1em; background-color: #8A1538; padding: 4px 10px; border-radius: 4px; box-shadow: 0 2px 5px rgba(138, 21, 56, 0.4); margin: 0 3px; display: inline-block; position: relative; max-width: 100%; overflow-wrap: break-word;">
                 <span class="column-f-label" style="position: absolute; top: -10px; left: 0; background-color: #FFD700; color: #8A1538; font-size: 9px; padding: 1px 4px; border-radius: 3px; font-weight: bold;">CỘT F</span>
                 ${formattedCell}
               </span>`;
    }
    
    return `<div style="margin: 10px 0;">${result}</div>`;
}

// Hiển thị thông báo
function showNotification(message, type = 'info') {
    // Tạo thông báo
    const notification = document.createElement('div');
    notification.className = 'notification ' + type;
    notification.textContent = message;
    
    // Style cho thông báo
    notification.style.position = 'fixed';
    notification.style.top = '20px';
    notification.style.left = '50%';
    notification.style.transform = 'translateX(-50%)';
    notification.style.padding = '10px 20px';
    notification.style.borderRadius = '5px';
    notification.style.zIndex = '1000';
    notification.style.boxShadow = '0 3px 6px rgba(0,0,0,0.2)';
    
    // Màu sắc dựa trên loại thông báo
    if (type === 'warning') {
        notification.style.backgroundColor = '#fff3cd';
        notification.style.color = '#856404';
        notification.style.border = '1px solid #ffeeba';
    } else if (type === 'error') {
        notification.style.backgroundColor = '#f8d7da';
        notification.style.color = '#721c24';
        notification.style.border = '1px solid #f5c6cb';
    } else {
        notification.style.backgroundColor = '#d4edda';
        notification.style.color = '#155724';
        notification.style.border = '1px solid #c3e6cb';
    }
    
    // Thêm vào body
    document.body.appendChild(notification);
    
    // Xóa thông báo sau 3 giây
    setTimeout(() => {
        notification.style.opacity = '0';
        notification.style.transition = 'opacity 0.5s';
        setTimeout(() => document.body.removeChild(notification), 500);
    }, 3000);
}

// Tải dữ liệu từ file Excel khi chọn file
function handleFileSelect(event) {
    const file = event.target.files[0];
    
    if (!file) return;
    
    // Cập nhật tên file được chọn
    document.getElementById('fileName').textContent = file.name;
    
    const reader = new FileReader();
    
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
            excelData = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });
            console.log('Đã tải dữ liệu Excel:', excelData.length, 'dòng');
            showNotification(`Đã tải thành công file ${file.name} với ${excelData.length} dòng dữ liệu.`);
        } catch (error) {
            console.error('Lỗi khi xử lý file:', error);
            showNotification('Không thể xử lý file Excel. Vui lòng kiểm tra định dạng file.', 'error');
        }
    };
    
    reader.onerror = function() {
        console.error('Lỗi khi đọc file');
        showNotification('Không thể đọc file. Vui lòng thử lại.', 'error');
    };
    
    reader.readAsArrayBuffer(file);
}

// Hàm xử lý ảnh và thực hiện OCR
async function processImageForOCR(imageDataUrl) {
    console.log('Processing image for OCR...');
    showNotification('Đang nhận dạng văn bản từ ảnh...');
    let text = '';
    try {
        text = await new Promise((resolve, reject) => {
            const img = new Image();

            img.onload = function() {
                const canvas = document.createElement('canvas');
                const ctx = canvas.getContext('2d');

                // Giới hạn kích thước tối đa
                let width = img.width;
                let height = img.height;
                const MAX_SIZE = 2048;
                if (width > MAX_SIZE || height > MAX_SIZE) {
                    if (width > height) {
                        height = Math.round((height * MAX_SIZE) / width);
                        width = MAX_SIZE;
                    } else {
                        width = Math.round((width * MAX_SIZE) / height);
                        height = MAX_SIZE;
                    }
                }
                canvas.width = width;
                canvas.height = height;
                ctx.drawImage(img, 0, 0, width, height);

                // Lấy dữ liệu ảnh
                let imageData = ctx.getImageData(0, 0, width, height);
                let data = imageData.data;

                // 1. Chuyển sang grayscale
                for (let i = 0; i < data.length; i += 4) {
                    const avg = (data[i] + data[i + 1] + data[i + 2]) / 3;
                    data[i] = data[i + 1] = data[i + 2] = avg;
                }

                // 2. Tăng tương phản mạnh hơn nữa
                const contrast = 4.0; // tăng giá trị này nếu cần
                const factor = (259 * (contrast * 255 + 255)) / (255 * (259 - contrast * 255));
                for (let i = 0; i < data.length; i += 4) {
                    let v = data[i];
                    v = factor * (v - 128) + 128;
                    v = Math.max(0, Math.min(255, v));
                    data[i] = data[i + 1] = data[i + 2] = v;
                }

                // 3. Binarization (ngưỡng trắng đen)
                const threshold = 125; // giảm xuống cho chữ mảnh, nền sáng
                for (let i = 0; i < data.length; i += 4) {
                    const value = data[i] > threshold ? 255 : 0;
                    data[i] = data[i + 1] = data[i + 2] = value;
                }

                // 4. Làm sắc nét nhẹ (sharpen)
                function sharpen(imageData, width, height) {
                    const weights = [0, -1, 0, -1, 5, -1, 0, -1, 0];
                    const side = Math.round(Math.sqrt(weights.length));
                    const halfSide = Math.floor(side / 2);
                    const src = imageData.data.slice();
                    for (let y = 0; y < height; y++) {
                        for (let x = 0; x < width; x++) {
                            let r = 0, g = 0, b = 0;
                            for (let cy = 0; cy < side; cy++) {
                                for (let cx = 0; cx < side; cx++) {
                                    const scy = y + cy - halfSide;
                                    const scx = x + cx - halfSide;
                                    if (scy >= 0 && scy < height && scx >= 0 && scx < width) {
                                        const srcOffset = (scy * width + scx) * 4;
                                        const wt = weights[cy * side + cx];
                                        r += src[srcOffset] * wt;
                                        g += src[srcOffset + 1] * wt;
                                        b += src[srcOffset + 2] * wt;
                                    }
                                }
                            }
                            const dstOffset = (y * width + x) * 4;
                            data[dstOffset] = Math.min(255, Math.max(0, r));
                            data[dstOffset + 1] = Math.min(255, Math.max(0, g));
                            data[dstOffset + 2] = Math.min(255, Math.max(0, b));
                        }
                    }
                }
                sharpen(imageData, width, height);

                ctx.putImageData(imageData, 0, 0);

                const processedImageDataUrl = canvas.toDataURL('image/jpeg', 1.0);

                // Thực hiện OCR với ảnh đã xử lý
                Tesseract.recognize(
                    processedImageDataUrl,
                    'vie',
                    {
                        logger: m => console.log(m),
                        tessedit_char_whitelist: '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyzÀÁÂÃÈÉÊÌÍÒÓÔÕÙÚĂĐĨŨƠàáâãèéêìòóôõùúăđĩũơƯĂẠẢẤẦẨẪẬẮẰẲẴẶẸẺẼỀỀỂẾưăạảấầẩẫậắằẳẵặẹẻẽềềểếỄỆỈỊỌỎỐỒỔỖỘỚỜỞỠỢỤỦỨỪễệỉịọỏốồổỗộớờởỡợụủứừỬỮỰỲỴÝỶỸửữựỳỵýỷỹ.,:-/ ',
                        preserve_interword_spaces: '1'
                    }
                ).then(({ data }) => {
                    let recognizedText = '';
                    if (data && data.text) {
                        recognizedText = data.text.trim();
                    }
                    showNotification('Đã nhận dạng văn bản từ ảnh');
                    resolve(recognizedText);
                }).catch(err => {
                    showNotification('Lỗi khi nhận dạng văn bản từ ảnh.', 'error');
                    reject(err);
                });
            };

            img.onerror = function() {
                showNotification('Không thể tải ảnh để xử lý.', 'error');
                reject(new Error('Could not load image'));
            };

            img.src = imageDataUrl;
        });
    } catch (e) {
        text = '';
    }
    // Nếu kết quả quá ngắn, thử lại bằng OCR.Space
    if (!text || text.length < 5) {
        text = await ocrSpaceRecognize(imageDataUrl);
    }
    return text;
}

// Hàm kiểm tra trình duyệt
function checkBrowser() {
    const userAgent = navigator.userAgent.toLowerCase();
    const isIOS = /iphone|ipad|ipod/.test(userAgent);
    const isChrome = /chrome/.test(userAgent);
    const isAndroid = /android/.test(userAgent);
    
    // Chỉ cảnh báo khi là Chrome trên iOS
    if (isIOS && isChrome && !isAndroid) {
        showNotification('Vui lòng sử dụng Safari để có trải nghiệm tốt nhất. Chrome trên iOS có thể không hỗ trợ đầy đủ tính năng chụp ảnh.', 'warning');
        return false;
    }
    return true;
}

// Hàm xử lý ảnh được chụp trực tiếp
async function handleCaptureSelect(event) {
    console.log('Bắt đầu xử lý ảnh chụp...');
    
    // Kiểm tra trình duyệt trước khi xử lý
    if (!checkBrowser()) {
        document.getElementById('captureName').textContent = 'Chưa chụp';
        return;
    }
    
    const file = event.target.files[0];

    if (!file) {
        console.log('Không có file được chọn');
        document.getElementById('captureName').textContent = 'Chưa chụp';
        return;
    }

    document.getElementById('captureName').textContent = 'Đã chụp ảnh';
    showNotification('Đã chụp ảnh. Vui lòng chọn vùng cần xử lý...');

    const reader = new FileReader();

    reader.onload = function(e) {
        console.log('Đã đọc xong file ảnh');
        const imageDataUrl = e.target.result;
        showImageCropModal(imageDataUrl);
    };

    reader.onerror = function() {
        console.error('Lỗi khi đọc ảnh chụp');
        showNotification('Không thể đọc ảnh chụp. Vui lòng thử lại.', 'error');
    };

    reader.readAsDataURL(file);
}

// Hàm xử lý việc chọn file ảnh để OCR
async function handleImageSelect(event) {
    // Kiểm tra trình duyệt trước khi xử lý
    if (!checkBrowser()) {
        document.getElementById('imageName').textContent = 'Chưa có ảnh';
        return;
    }
    
    const file = event.target.files[0];
    
    if (!file) {
        document.getElementById('imageName').textContent = 'Chưa có ảnh';
        return;
    }
    
    // Cập nhật tên file ảnh được chọn
    document.getElementById('imageName').textContent = file.name;
    showNotification('Đã chọn ảnh. Vui lòng chọn vùng cần xử lý...');
    
    const reader = new FileReader();
    
    reader.onload = function(e) {
        const imageDataUrl = e.target.result;
        showImageCropModal(imageDataUrl);
    };
    
    reader.onerror = function() {
        console.error('Lỗi khi đọc file ảnh');
        showNotification('Không thể đọc file ảnh. Vui lòng thử lại.', 'error');
    };
    
    reader.readAsDataURL(file);
}

// Hàm hiển thị modal chọn vùng ảnh
function showImageCropModal(imageDataUrl) {
    console.log('Bắt đầu hiển thị modal...');
    const modal = document.getElementById('imageCropModal');
    const cropImage = document.getElementById('cropImage');
    
    if (!modal || !cropImage) {
        console.error('Không tìm thấy modal hoặc cropImage element');
        return;
    }
    
    // Đặt src cho ảnh
    cropImage.src = imageDataUrl;
    currentImage = cropImage;
    
    // Hiển thị modal
    modal.style.display = 'block';
    console.log('Modal đã được hiển thị');
    
    // Xóa vùng chọn cũ nếu có
    if (selectionOverlay) {
        selectionOverlay.remove();
        selectionOverlay = null;
    }
    
    // Thêm sự kiện cho việc chọn vùng
    cropImage.onload = function() {
        console.log('Ảnh đã được tải xong');
        const container = document.querySelector('.image-container');
        
        if (!container) {
            console.error('Không tìm thấy image-container');
            return;
        }
        
        // Xóa các event listener cũ
        container.removeEventListener('mousedown', handleMouseDown);
        container.removeEventListener('mousemove', handleMouseMove);
        container.removeEventListener('mouseup', handleMouseUp);
        container.removeEventListener('touchstart', handleTouchStart);
        container.removeEventListener('touchmove', handleTouchMove);
        container.removeEventListener('touchend', handleTouchEnd);
        
        // Thêm sự kiện mới
        container.addEventListener('mousedown', handleMouseDown);
        container.addEventListener('mousemove', handleMouseMove);
        container.addEventListener('mouseup', handleMouseUp);
        container.addEventListener('touchstart', handleTouchStart, { passive: false });
        container.addEventListener('touchmove', handleTouchMove, { passive: false });
        container.addEventListener('touchend', handleTouchEnd);
        
        console.log('Đã thêm các event listener cho container');
    };
    
    cropImage.onerror = function() {
        console.error('Lỗi khi tải ảnh');
        showNotification('Không thể tải ảnh. Vui lòng thử lại.', 'error');
    };
}

// Hàm xử lý sự kiện mousedown
function handleMouseDown(e) {
    const container = document.querySelector('.image-container');
    const rect = container.getBoundingClientRect();
    
    // Kiểm tra nếu click vào handle để resize
    if (e.target.classList.contains('selection-handle')) {
        isResizing = true;
        currentHandle = e.target;
        startX = e.clientX;
        startY = e.clientY;
        originalX = parseInt(selectionOverlay.style.left);
        originalY = parseInt(selectionOverlay.style.top);
        originalWidth = parseInt(selectionOverlay.style.width);
        originalHeight = parseInt(selectionOverlay.style.height);
        return;
    }
    
    // Kiểm tra nếu click vào vùng chọn để di chuyển
    if (e.target === selectionOverlay) {
        isMoving = true;
        startX = e.clientX;
        startY = e.clientY;
        originalX = parseInt(selectionOverlay.style.left);
        originalY = parseInt(selectionOverlay.style.top);
        return;
    }
    
    // Tạo vùng chọn mới
    isSelecting = true;
    startX = e.clientX - rect.left;
    startY = e.clientY - rect.top;
    
    // Xóa vùng chọn cũ nếu có
    if (selectionOverlay) {
        selectionOverlay.remove();
    }
    
    // Tạo overlay mới
    selectionOverlay = document.createElement('div');
    selectionOverlay.className = 'selection-overlay';
    selectionOverlay.style.left = startX + 'px';
    selectionOverlay.style.top = startY + 'px';
    
    // Thêm các handle để resize
    const handles = ['nw', 'ne', 'sw', 'se'];
    handles.forEach(pos => {
        const handle = document.createElement('div');
        handle.className = `selection-handle ${pos}`;
        selectionOverlay.appendChild(handle);
    });
    
    container.appendChild(selectionOverlay);
}

// Hàm xử lý sự kiện mousemove
function handleMouseMove(e) {
    const container = document.querySelector('.image-container');
    const rect = container.getBoundingClientRect();
    
    if (isSelecting) {
        const currentX = e.clientX - rect.left;
        const currentY = e.clientY - rect.top;
        
        const width = Math.abs(currentX - startX);
        const height = Math.abs(currentY - startY);
        
        selectionOverlay.style.width = width + 'px';
        selectionOverlay.style.height = height + 'px';
        selectionOverlay.style.left = Math.min(startX, currentX) + 'px';
        selectionOverlay.style.top = Math.min(startY, currentY) + 'px';
    }
    else if (isMoving) {
        const deltaX = e.clientX - startX;
        const deltaY = e.clientY - startY;
        
        const newX = originalX + deltaX;
        const newY = originalY + deltaY;
        
        // Giới hạn vùng chọn trong container
        const maxX = rect.width - parseInt(selectionOverlay.style.width);
        const maxY = rect.height - parseInt(selectionOverlay.style.height);
        
        selectionOverlay.style.left = Math.max(0, Math.min(newX, maxX)) + 'px';
        selectionOverlay.style.top = Math.max(0, Math.min(newY, maxY)) + 'px';
    }
    else if (isResizing) {
        const deltaX = e.clientX - startX;
        const deltaY = e.clientY - startY;
        
        switch(currentHandle.className) {
            case 'selection-handle nw':
                selectionOverlay.style.width = (originalWidth - deltaX) + 'px';
                selectionOverlay.style.height = (originalHeight - deltaY) + 'px';
                selectionOverlay.style.left = (originalX + deltaX) + 'px';
                selectionOverlay.style.top = (originalY + deltaY) + 'px';
                break;
            case 'selection-handle ne':
                selectionOverlay.style.width = (originalWidth + deltaX) + 'px';
                selectionOverlay.style.height = (originalHeight - deltaY) + 'px';
                selectionOverlay.style.top = (originalY + deltaY) + 'px';
                break;
            case 'selection-handle sw':
                selectionOverlay.style.width = (originalWidth - deltaX) + 'px';
                selectionOverlay.style.height = (originalHeight + deltaY) + 'px';
                selectionOverlay.style.left = (originalX + deltaX) + 'px';
                break;
            case 'selection-handle se':
                selectionOverlay.style.width = (originalWidth + deltaX) + 'px';
                selectionOverlay.style.height = (originalHeight + deltaY) + 'px';
                break;
        }
        
        // Đảm bảo kích thước tối thiểu
        const minSize = 20;
        if (parseInt(selectionOverlay.style.width) < minSize) {
            selectionOverlay.style.width = minSize + 'px';
        }
        if (parseInt(selectionOverlay.style.height) < minSize) {
            selectionOverlay.style.height = minSize + 'px';
        }
    }
}

// Hàm xử lý sự kiện mouseup
function handleMouseUp() {
    isSelecting = false;
    isMoving = false;
    isResizing = false;
    currentHandle = null;
}

// Hàm xử lý sự kiện touchstart
function handleTouchStart(e) {
    console.log('Touch start event');
    e.preventDefault(); // Ngăn chặn scroll khi touch
    const touch = e.touches[0];
    const container = document.querySelector('.image-container');
    const rect = container.getBoundingClientRect();
    
    // Kiểm tra nếu touch vào handle để resize
    if (e.target.classList.contains('selection-handle')) {
        console.log('Touch vào handle để resize');
        isResizing = true;
        currentHandle = e.target;
        startX = touch.clientX;
        startY = touch.clientY;
        originalX = parseInt(selectionOverlay.style.left);
        originalY = parseInt(selectionOverlay.style.top);
        originalWidth = parseInt(selectionOverlay.style.width);
        originalHeight = parseInt(selectionOverlay.style.height);
        return;
    }
    
    // Kiểm tra nếu touch vào vùng chọn để di chuyển
    if (e.target === selectionOverlay) {
        console.log('Touch vào vùng chọn để di chuyển');
        isMoving = true;
        startX = touch.clientX;
        startY = touch.clientY;
        originalX = parseInt(selectionOverlay.style.left);
        originalY = parseInt(selectionOverlay.style.top);
        return;
    }
    
    // Tạo vùng chọn mới
    console.log('Bắt đầu tạo vùng chọn mới');
    isSelecting = true;
    startX = touch.clientX - rect.left;
    startY = touch.clientY - rect.top;
    
    // Xóa vùng chọn cũ nếu có
    if (selectionOverlay) {
        selectionOverlay.remove();
        selectionOverlay = null;
    }
    
    // Tạo overlay mới
    selectionOverlay = document.createElement('div');
    selectionOverlay.className = 'selection-overlay';
    selectionOverlay.style.left = startX + 'px';
    selectionOverlay.style.top = startY + 'px';
    
    // Thêm các handle để resize
    const handles = ['nw', 'ne', 'sw', 'se'];
    handles.forEach(pos => {
        const handle = document.createElement('div');
        handle.className = `selection-handle ${pos}`;
        selectionOverlay.appendChild(handle);
    });
    
    container.appendChild(selectionOverlay);
    console.log('Đã tạo xong vùng chọn mới');
}

// Hàm xử lý sự kiện touchmove
function handleTouchMove(e) {
    e.preventDefault(); // Ngăn chặn scroll khi touch
    const touch = e.touches[0];
    const container = document.querySelector('.image-container');
    const rect = container.getBoundingClientRect();
    
    if (isSelecting) {
        const currentX = touch.clientX - rect.left;
        const currentY = touch.clientY - rect.top;
        
        const width = Math.abs(currentX - startX);
        const height = Math.abs(currentY - startY);
        
        selectionOverlay.style.width = width + 'px';
        selectionOverlay.style.height = height + 'px';
        selectionOverlay.style.left = Math.min(startX, currentX) + 'px';
        selectionOverlay.style.top = Math.min(startY, currentY) + 'px';
    }
    else if (isMoving) {
        const deltaX = touch.clientX - startX;
        const deltaY = touch.clientY - startY;
        
        const newX = originalX + deltaX;
        const newY = originalY + deltaY;
        
        // Giới hạn vùng chọn trong container
        const maxX = rect.width - parseInt(selectionOverlay.style.width);
        const maxY = rect.height - parseInt(selectionOverlay.style.height);
        
        selectionOverlay.style.left = Math.max(0, Math.min(newX, maxX)) + 'px';
        selectionOverlay.style.top = Math.max(0, Math.min(newY, maxY)) + 'px';
    }
    else if (isResizing) {
        const deltaX = touch.clientX - startX;
        const deltaY = touch.clientY - startY;
        
        switch(currentHandle.className) {
            case 'selection-handle nw':
                selectionOverlay.style.width = (originalWidth - deltaX) + 'px';
                selectionOverlay.style.height = (originalHeight - deltaY) + 'px';
                selectionOverlay.style.left = (originalX + deltaX) + 'px';
                selectionOverlay.style.top = (originalY + deltaY) + 'px';
                break;
            case 'selection-handle ne':
                selectionOverlay.style.width = (originalWidth + deltaX) + 'px';
                selectionOverlay.style.height = (originalHeight - deltaY) + 'px';
                selectionOverlay.style.top = (originalY + deltaY) + 'px';
                break;
            case 'selection-handle sw':
                selectionOverlay.style.width = (originalWidth - deltaX) + 'px';
                selectionOverlay.style.height = (originalHeight + deltaY) + 'px';
                selectionOverlay.style.left = (originalX + deltaX) + 'px';
                break;
            case 'selection-handle se':
                selectionOverlay.style.width = (originalWidth + deltaX) + 'px';
                selectionOverlay.style.height = (originalHeight + deltaY) + 'px';
                break;
        }
        
        // Đảm bảo kích thước tối thiểu
        const minSize = 20;
        if (parseInt(selectionOverlay.style.width) < minSize) {
            selectionOverlay.style.width = minSize + 'px';
        }
        if (parseInt(selectionOverlay.style.height) < minSize) {
            selectionOverlay.style.height = minSize + 'px';
        }
    }
}

// Hàm xử lý sự kiện touchend
function handleTouchEnd() {
    isSelecting = false;
    isMoving = false;
    isResizing = false;
    currentHandle = null;
}

// Hàm cắt ảnh theo vùng đã chọn
async function cropImage() {
    if (!selectionOverlay || !currentImage) return;
    
    const canvas = document.createElement('canvas');
    const ctx = canvas.getContext('2d');
    
    const rect = selectionOverlay.getBoundingClientRect();
    const imageRect = currentImage.getBoundingClientRect();
    
    // Tính toán tỷ lệ giữa ảnh gốc và ảnh hiển thị
    const scaleX = currentImage.naturalWidth / imageRect.width;
    const scaleY = currentImage.naturalHeight / imageRect.height;
    
    // Tính toán vị trí và kích thước thực tế của vùng cắt
    const cropX = (rect.left - imageRect.left) * scaleX;
    const cropY = (rect.top - imageRect.top) * scaleY;
    const cropWidth = rect.width * scaleX;
    const cropHeight = rect.height * scaleY;
    
    canvas.width = cropWidth;
    canvas.height = cropHeight;
    
    // Vẽ phần ảnh đã cắt
    ctx.drawImage(
        currentImage,
        cropX, cropY, cropWidth, cropHeight,
        0, 0, cropWidth, cropHeight
    );
    
    // Chuyển đổi canvas thành Data URL
    const croppedImageDataUrl = canvas.toDataURL('image/jpeg');
    
    // Đóng modal
    document.getElementById('imageCropModal').style.display = 'none';
    
    // Xử lý ảnh đã cắt
    try {
        const recognizedText = await processImageForOCR(croppedImageDataUrl);
        // Thay thế các ký tự xuống dòng bằng khoảng trắng và trim
        const cleanedText = recognizedText.replace(/[\r\n]+/g, ' ').trim();
        // Điền văn bản đã xử lý vào ô tìm kiếm
        document.getElementById('searchInput').value = cleanedText;
        showNotification('Đã nhận dạng và điền văn bản vào ô tìm kiếm');
        // Tự động gọi tìm kiếm sau khi nhận dạng thành công
        searchExcelData(cleanedText);
        
    } catch (error) {
        console.error('Lỗi xử lý ảnh cắt:', error);
        showNotification('Lỗi khi xử lý ảnh cắt.', 'error');
    }
}

// Thêm kiểm tra trình duyệt khi trang được tải
window.onload = () => {
    // Kiểm tra trình duyệt
    checkBrowser();
    
    // Đăng ký sự kiện cho nút chọn file Excel
    document.getElementById('fileInput').addEventListener('change', handleFileSelect);
    
    // Đăng ký sự kiện cho nút chọn file ảnh
    document.getElementById('imageInput').addEventListener('change', handleImageSelect);
    
    // Thêm trình lắng nghe sự kiện paste
    document.body.addEventListener('paste', handlePaste);
    
    // Thêm xử lý paste cho ô tìm kiếm
    document.getElementById('searchInput').addEventListener('paste', handleSearchInputPaste);
    
    // Đăng ký sự kiện cho nút chụp ảnh
    document.getElementById('captureInput').addEventListener('change', handleCaptureSelect);
    
    // Tải file Excel mặc định
    loadDefaultExcelFile();
    
    // Khởi tạo nhận diện giọng nói
    initSpeechRecognition();
    
    // Đăng ký sự kiện cho các nút khác
    document.getElementById('startListeningButton').onclick = startListening;
    
    document.getElementById('clearResultsButton').onclick = clearResults;
    
    // Gọi hàm tìm kiếm khi nhấn nút tìm kiếm
    document.getElementById('searchButton').onclick = () => {
        const searchTerm = document.getElementById('searchInput').value;
        searchExcelData(searchTerm);
    };
    
    // Tìm kiếm khi nhấn phím Enter
    document.getElementById('searchInput').addEventListener('keypress', function(event) {
        if (event.key === 'Enter') {
            event.preventDefault();
            const searchTerm = document.getElementById('searchInput').value;
            searchExcelData(searchTerm);
        }
    });
    
    // Thêm event listeners cho modal
    document.querySelector('.close').onclick = function() {
        document.getElementById('imageCropModal').style.display = 'none';
    };
    
    document.getElementById('cropButton').onclick = cropImage;
    
    document.getElementById('cancelCropButton').onclick = function() {
        document.getElementById('imageCropModal').style.display = 'none';
    };
    
    // Đóng modal khi click bên ngoài
    window.onclick = function(event) {
        const modal = document.getElementById('imageCropModal');
        if (event.target == modal) {
            modal.style.display = 'none';
        }
    };
};

function displayResults(data) {
    const resultArea = document.getElementById("resultArea");
    resultArea.innerHTML = "";

    data.forEach((row, rowIndex) => {
        if (rowIndex === 0) return; // Bỏ qua header

        const div = document.createElement("div");
        let rowContent = "";

        row.forEach((cell, colIndex) => {
            // Nếu là cột F (index 5), bọc trong span với class nổi bật
            if (colIndex === 5) {
                rowContent += `<span class="column-f-highlight"><span class="column-f-label">Cột F</span>${cell}</span> `;
            } else {
                rowContent += `<span>${cell}</span> `;
            }
        });

        div.innerHTML = rowContent.trim();
        resultArea.appendChild(div);
    });
}

// Tải file Excel mặc định khi trang web được tải
function loadDefaultExcelFile() {
    try {
        const xhr = new XMLHttpRequest();
        xhr.open('GET', 'data.xlsx', true);
        xhr.responseType = 'arraybuffer';

        xhr.onload = function () {
            if (xhr.status === 200) {
                const data = new Uint8Array(xhr.response);
                const workbook = XLSX.read(data, { type: 'array' });
                const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                excelData = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });
                console.log('Đã tải dữ liệu Excel mặc định:', excelData.length, 'dòng');
                document.getElementById('fileName').textContent = 'data.xlsx';
                showNotification(`Đã tải thành công file mặc định với ${excelData.length} dòng dữ liệu.`);
            } else {
                console.warn('Không thể tải file Excel mặc định:', xhr.status);
                showNotification('Không tìm thấy file mặc định. Vui lòng chọn file Excel.', 'warning');
            }
        };

        xhr.onerror = function() {
            console.warn('Lỗi khi tải file Excel mặc định');
            showNotification('Không tìm thấy file mặc định. Vui lòng chọn file Excel.', 'warning');
        };

        xhr.send();
    } catch (error) {
        console.warn('Lỗi khi tải file Excel mặc định:', error);
    }
}

// Hàm xử lý sự kiện paste
function handlePaste(event) {
    const items = (event.clipboardData || event.originalEvent.clipboardData).items;
    let blob = null;

    for (const item of items) {
        // Tìm kiếm item có kiểu là image
        if (item.type.indexOf('image') === 0) {
            blob = item.getAsFile();
            break;
        }
    }

    if (blob) {
        console.log('Đã dán ảnh từ clipboard');
        showNotification('Đã phát hiện ảnh từ clipboard. Đang xử lý...');
        
        // Đọc blob ảnh dưới dạng Data URL
        const reader = new FileReader();
        reader.onload = function(e) {
            const imageDataUrl = e.target.result;
            // Gọi hàm xử lý OCR chung
            processImageForOCR(imageDataUrl);
        };
        reader.onerror = function() {
            console.error('Lỗi khi đọc blob ảnh từ clipboard');
            showNotification('Không thể đọc ảnh từ clipboard.', 'error');
        };
        reader.readAsDataURL(blob);
    } else {
        console.log('Không có ảnh trong clipboard');
    }
}

// Hàm xử lý paste cho ô tìm kiếm
async function handleSearchInputPaste(e) {
    console.log('Paste event triggered on search input');
    const items = (e.clipboardData || window.clipboardData).items;
    let blob = null;

    for (const item of items) {
        // Tìm kiếm item có kiểu là image
        if (item.type.indexOf('image') === 0) {
            blob = item.getAsFile();
            console.log('Image found in clipboard');
            break;
        }
    }

    if (blob) {
        e.preventDefault(); // Ngăn chặn paste mặc định của ảnh
        console.log('Processing pasted image...');
        showNotification('Đã phát hiện ảnh. Đang nhận dạng văn bản...');
        
        // Đọc blob ảnh dưới dạng Data URL
        const reader = new FileReader();
        reader.onload = async function(e) {
            const imageDataUrl = e.target.result;
            try {
                // Xử lý ảnh đầu vào giống như nút chụp ảnh/chọn ảnh
                const recognizedText = await processImageForOCR(imageDataUrl);
                // Thay thế các ký tự xuống dòng bằng khoảng trắng và trim
                const cleanedText = recognizedText.replace(/[\r\n]+/g, ' ').trim();
                // Điền văn bản đã xử lý vào ô tìm kiếm
                document.getElementById('searchInput').value = cleanedText;
                showNotification('Đã nhận dạng và điền văn bản vào ô tìm kiếm');
                // Tự động gọi tìm kiếm sau khi nhận dạng thành công
                searchExcelData(cleanedText);
                
            } catch (error) {
                console.error('Lỗi xử lý ảnh dán:', error);
                showNotification('Lỗi khi xử lý ảnh dán.', 'error');
            }
        };
        reader.onerror = function() {
            console.error('Lỗi khi đọc ảnh từ clipboard');
            showNotification('Không thể đọc ảnh từ clipboard.', 'error');
        };
        reader.readAsDataURL(blob);
    } else {
        console.log('No image found, processing as text paste');
        // Nếu không phải ảnh, xử lý paste văn bản bình thường
        const pastedText = (e.clipboardData || window.clipboardData).getData('text');
        const cleanedText = pastedText.replace(/[\r\n]+/g, ' ').trim();

        e.preventDefault(); // THÊM DÒNG NÀY để ngăn hành vi paste mặc định

        e.target.value = cleanedText; // Ghi đè giá trị input
        console.log('Pasted text:', cleanedText);

        // Tự động gọi tìm kiếm sau khi paste văn bản
        searchExcelData(cleanedText);
    }
}
async function ocrSpaceRecognize(imageDataUrl) {
    showNotification('Đang gửi ảnh lên OCR.Space...');
    const apiKey = 'helloworld'; // API key miễn phí mặc định của OCR.Space

    // OCR.Space chỉ nhận base64 không có tiền tố "data:image/jpeg;base64,"
    const base64Image = imageDataUrl.replace(/^data:image\/(png|jpeg);base64,/, '');

    const formData = new FormData();
    formData.append('base64Image', 'data:image/jpeg;base64,' + base64Image);
    formData.append('language', 'vie');
    formData.append('isOverlayRequired', 'false');

    try {
        const response = await fetch('https://api.ocr.space/parse/image', {
            method: 'POST',
            headers: {
                apikey: apiKey
            },
            body: formData
        });
        const result = await response.json();
        if (result.IsErroredOnProcessing) {
            showNotification('OCR.Space lỗi: ' + result.ErrorMessage, 'error');
            return '';
        }
        const text = result.ParsedResults && result.ParsedResults[0] ? result.ParsedResults[0].ParsedText : '';
        showNotification('Đã nhận dạng văn bản từ ảnh (OCR.Space)');
        return text.trim();
    } catch (error) {
        showNotification('Lỗi khi gửi ảnh lên OCR.Space.', 'error');
        return '';
    }
}