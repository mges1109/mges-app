const express = require('express');
const puppeteer = require('puppeteer');
const path = require('path');
const XLSX = require('xlsx');  // Thêm thư viện XLSX để tạo Excel
const cheerio = require('cheerio'); // Thêm cheerio vào dự án
const { exec } = require('child_process');
const axios = require('axios');
const fs = require('fs');


const app = express();
const port = 3001;
app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(express.static(path.join(__dirname, 'public')));
app.use(express.static(path.join(__dirname, 'public/web')));

function docSoThanhChu(so, options = {}) {
  let dau = '';
  if (so < 0) {
    dau = 'âm ';
    so = -so;
  } else if (so === 0) {
    return 'không đồng';
  }
  
  let phanNguyen = Math.floor(so);
  let phanThapPhan = Math.round((so - phanNguyen) * 100);

  let chuoi = docPhanNguyen(phanNguyen);

  if (phanThapPhan > 0) {
    chuoi += ' phẩy ' + docPhanNguyen(phanThapPhan) + ' đồng';
  } else {
    chuoi += ' đồng';
  }

  chuoi = dau + chuoi;
  if (options.capitalize) {
    chuoi = chuoi.charAt(0).toUpperCase() + chuoi.slice(1);
  }
  if (options.comma) {
    chuoi = chuoi.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  }
  return chuoi;
}

function docPhanNguyen(so) {
  const hangNgan = ['', 'nghìn', 'triệu', 'tỷ'];
  if (so === 0) return '';
  let chuoi = '';
  let hang = 0;
  while (so > 0) {
    let hangDonVi = so % 1000;
    if (hangDonVi > 0) {
      chuoi = docHang(hangDonVi) + ' ' + hangNgan[hang] + ' ' + chuoi;
    }
    so = Math.floor(so / 1000);
    hang++;
  }
  return chuoi.trim();
}

function docHang(so) {
  const hangTram = ['', 'một trăm', 'hai trăm', 'ba trăm', 'bốn trăm', 'năm trăm', 'sáu trăm', 'bảy trăm', 'tám trăm', 'chín trăm'];
  const donVi = ['', 'một', 'hai', 'ba', 'bốn', 'năm', 'sáu', 'bảy', 'tám', 'chín'];
  const hangChuc = ['', 'mười', 'hai mươi', 'ba mươi', 'bốn mươi', 'năm mươi', 'sáu mươi', 'bảy mươi', 'tám mươi', 'chín mươi'];
  
  let chuoi = '';
  let tram = Math.floor(so / 100);
  let chuc = Math.floor((so % 100) / 10);
  let donvi = so % 10;

  chuoi += hangTram[tram];
  if (tram > 0 && (chuc > 0 || donvi > 0)) chuoi += ' ';
  chuoi += hangChuc[chuc];
  if (chuc > 1 && donvi > 0) chuoi += ' ';
  chuoi += donVi[donvi];

  return chuoi;
}

app.post('/print', (req, res) => {
  const usernameId = req.query.key; // Lấy API key từ query string
  const printId = req.query.printid; // Lấy printerId từ query string
  const filename = req.body.filename; // Nhận tên file PDF từ body

  if (!filename || !usernameId || !printId) {
      
      return res.status(400).json({ error: 'Thiếu filename, API key hoặc printerId' });
  }

  // Đường dẫn đầy đủ đến file PDF
  const pdfFilePath = path.join(__dirname, 'public', filename);


  // Đọc file PDF và mã hóa thành base64
  const pdfBuffer = fs.readFileSync(pdfFilePath);
  const pdfBase64 = pdfBuffer.toString('base64');

  // Gửi lệnh in tới PrintNode
  axios.post('https://api.printnode.com/printjobs', {
      printerId: printId,
      title: 'Hóa đơn đơn hàng',
      contentType: 'pdf_base64',
      content: pdfBase64,
      source: 'Node.js App'
  }, {
      auth: {
          username: usernameId, 
          password: '' 
      }
  }).then(response => {
      console.log("Lệnh in đã được gửi:", response.data);
      res.json({ message: 'Lệnh in đã được gửi thành công' });
  }).catch(error => {
      console.error('Lỗi khi gửi lệnh in:', error);
      res.status(500).json({ error: 'Có lỗi xảy ra khi gửi lệnh in' });
  });
});




app.get('/printers', (req, res) => {
  const usernameId = req.query.key; // Lấy API key từ query string

  axios.get('https://api.printnode.com/printers', {
      auth: {
          username: `${usernameId}`, // Lấy API key từ query string
          password: '' // Không cần mật khẩu
      }
  }).then(response => {
      res.json(response.data);
  }).catch(error => {
      console.error('Lỗi khi lấy danh sách máy in:', error);
      res.status(500).json({ error: 'Không thể lấy danh sách máy in' });
  });
});


// Hàm trích xuất dữ liệu khách hàng từ HTML
function extractOrderDataFromHTML(html) {
  const $ = cheerio.load(html); // Sử dụng cheerio để tải HTML
  const orders = [];

  // Tìm tất cả các hàng trong bảng đơn hàng
  $('.loaddh tr').each((index, row) => {
    const columns = $(row).find('td'); // Tìm các ô dữ liệu trong từng hàng
    if (columns.length === 5) { // Đảm bảo có đủ 5 cột (STT, Sản phẩm, Số lượng, Đơn giá, Thành tiền)
      orders.push({
        sp: $(columns[1]).text().trim(),  // Sản phẩm
        sl: $(columns[2]).text().trim(),  // Số lượng
        dg: $(columns[3]).text().trim(),  // Đơn giá
        tt: $(columns[4]).text().trim()   // Thành tiền
      });
    }
  });

  return orders;
}

function extractCustomerDataFromHTML(html) {
  const $ = cheerio.load(html); // Sử dụng cheerio để phân tích cú pháp HTML
  const customerName = $('.tkh').text().trim();
  const customerPhone = $('.sdt').text().trim();
  const customerAddress = $('.dc').text().trim();

  return {
    ten_kh: customerName,
    sdt: customerPhone,
    dc: customerAddress,
  };
}
app.get('/download-excel/:id', async (req, res) => {
  const customerId = req.params.id;
  const url = `http://hoadon1.netlify.app/?id=${customerId}`;

  try {
    // Sử dụng Puppeteer để tải trang
    const browser = await puppeteer.launch();
    const page = await browser.newPage();
    await page.goto(url);
    
    // Chờ đến khi đơn hàng được tải đầy đủ
    await page.waitForFunction(() => document.querySelector('.loaddh').children.length > 0);
    
    // Lấy nội dung HTML sau khi trang đã tải xong
    const html = await page.content();

    // Trích xuất dữ liệu khách hàng và đơn hàng từ HTML
    const customerData = extractCustomerDataFromHTML(html);
    const orderData = extractOrderDataFromHTML(html);

    // Kiểm tra nếu không có dữ liệu khách hàng hoặc đơn hàng
    if (!customerData.ten_kh || orderData.length === 0) {
      return res.status(400).send('Dữ liệu không hợp lệ');
    }

    // Chuẩn bị dữ liệu cho file Excel
    const worksheetData = [
      { A: 'Công ty TNHH ABC' },
      { A: 'Địa chỉ: 372 Cách mạng tháng 8, Quận 3, HCM' },
      { A: 'Email: ncq.hct1109@gmail.com' },
      { A: '' }, // Dòng trống để tạo khoảng cách
      { A: '' },
      { A: '' }, // Dòng trống để tạo khoảng cách
      { A: `Tên khách hàng: ${customerData.ten_kh}` },
      { A: `Số điện thoại: ${customerData.sdt}` },
      { A: `Địa chỉ: ${customerData.dc}` },
      { A: '' },
      { A: 'STT', B: 'Sản phẩm', C: 'Số lượng', D: 'Đơn giá', E: 'Thành tiền' },
    ];

    // Thêm dữ liệu đơn hàng vào
    orderData.forEach((order, index) => {
      worksheetData.push({
        A: index + 1,
        B: order.sp,
        C: order.sl,
        D: order.dg,
        E: order.tt
      });
    });

    // Tính tổng tiền, VAT và số tiền cần thanh toán
    const totalAmount = orderData.reduce((total, order) => {
      const amount = parseFloat(order.tt.replace(/\./g, '').replace(/[^0-9.-]+/g, "")); // Loại bỏ dấu chấm và ký tự không phải số
      return total + amount;
    }, 0);

    const vatAmount = (totalAmount * 0.10).toFixed(0); // 10% VAT
    const totalWithVAT = (parseFloat(totalAmount) + parseFloat(vatAmount)).toFixed(0); // Tổng cộng sau khi thêm VAT
    const amountToPayInWords = docSoThanhChu(totalWithVAT, { capitalize: true });

    // Định dạng số tiền bằng cách thêm dấu chấm cho hàng nghìn
    function formatCurrency(amount) {
      return parseInt(amount, 10).toLocaleString('vi-VN');
    }

    // Đẩy dữ liệu tổng tiền, VAT, và số tiền cần thanh toán vào Excel
    worksheetData.push(
      { A: '' },
      { A: `Tổng cộng: ${formatCurrency(totalAmount)} ₫` },
      { A: `VAT (10%): ${formatCurrency(vatAmount)} ₫` },
      { A: `Số tiền cần thanh toán: ${formatCurrency(totalWithVAT)} ₫` },
      { A: `Bằng chữ: ${amountToPayInWords}` }
    );

    // Tạo workbook và worksheet cho Excel
    const worksheet = XLSX.utils.json_to_sheet(worksheetData, { skipHeader: true });
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, `hoadon_${customerId}`);

    // Hợp nhất các cột B, C và D cho dòng "HÓA ĐƠN BÁN HÀNG"
    worksheet['!merges'] = [
      { s: { r: 3, c: 1 }, e: { r: 3, c: 3 } }, // Hợp nhất các ô từ B4 đến D4
    ];

    // Căn giữa nội dung của dòng HÓA ĐƠN BÁN HÀNG
    worksheet['B4'] = { v: 'HÓA ĐƠN BÁN HÀNG', s: { alignment: { horizontal: 'center', vertical: 'center' }, font: { bold: true } } };

    // Đường dẫn để lưu file Excel
    const excelFilePath = path.join(__dirname, `hoadon_${customerId}.xlsx`);
    XLSX.writeFile(workbook, excelFilePath);

    // Trả file Excel cho máy khách
    res.download(excelFilePath, `hoadon_${customerId}.xlsx`);

    // Đóng trình duyệt
    await browser.close();
  } catch (error) {
    console.error('Lỗi khi tạo file Excel:', error);
    res.status(500).send('Có lỗi xảy ra khi tạo file Excel');
  }
});



app.get('/data', async (req, res) => {
  const { id, link, printid, key, width = '69mm', height = '297mm' } = req.query;
  const url = `https://${link}?id=${id}`;
  const filename = `hoadon_${id}.pdf`;

  try {
    const browser = await puppeteer.launch();
    const page = await browser.newPage();
    await page.goto(url);
    await page.waitForFunction(() => document.querySelector('.loaddh').children.length > 0);

    // Đường dẫn đầy đủ đến file PDF
    const pdfFilePath = path.join(__dirname, 'public', filename);

    // Thiết lập khổ giấy tùy chỉnh
    const pdfOptions = {
      path: pdfFilePath,
      width: width,      // Chiều rộng của khổ giấy từ query string hoặc mặc định là 69mm
      height: height,   // Chiều dài của khổ giấy
      printBackground: true,
      margin: { top: '2cm', bottom: '2cm', left: '5mm', right: '5mm' },
    };

    await page.pdf(pdfOptions);
    await browser.close();

    res.send(`
      <!DOCTYPE html>
      <html lang="vi">
      <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>PDF Viewer</title>
        <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0-beta3/css/all.min.css">
        <style>
          body, html {
            height: 100%;
            margin: 0;
            width: 100%;
            display: flex;
            flex-direction: column;
            align-items: center;
            background-color: #5a5b5a;
            overflow-y: auto;
            padding-top: 22px; 
          }
          .toolbarButton {
              background-color: #333;
              color: #fff;
              border: none;
              padding: 5px 10px;
              margin-left: 5px;
              cursor: pointer;
              border-radius: 3px;
              display: flex;
              align-items: center;
          }
          #controls {
              position: fixed;
              top: 0;
              width: 100%;
              background-color: #333;
              color: #fff;
              padding: 15px;
              z-index: 1000;
              display: flex;
              justify-content: space-between; /* Điều chỉnh khoảng cách giữa các phần */
              align-items: center;
          }
          .toolbarLeft {
              flex-grow: 1; /* Đẩy nút "In Nhanh" về bên trái */
              display: flex;
              align-items: center;
          }

          #page-info {
              flex-grow: 50;
              text-align: center;
          }

          .toolbarRight {
              display: flex;
              align-items: center;
          }
          .toolbarButton i {
           
            margin: 3px 10px;
          }
          .pdf-page {
            border: 1px solid #000;
            box-shadow: 0px 0px 15px rgba(0, 0, 0, 0.1);
            margin: 10px 0;
            width: 99%;
            max-width: 400px;
            background-color: #fff;
          }
        </style>
      </head>
      <body>
        <div id="controls">
          <div class="toolbarLeft">
              <button id="trigger-print" class="toolbarButton">
                  <i class="fas fa-print"></i> In Nhanh
              </button>
              <button id="download-excel" class="toolbarButton">
                <i class="fas fa-file-excel"></i> Tải Excel
              </button>
          </div>
          <span id="page-info">1/1</span>
          <div class="toolbarRight">
              <button id="download" class="toolbarButton">
                  <i class="fas fa-download"></i>
              </button>
              <button id="print" class="toolbarButton">
                  <i class="fas fa-print"></i>
              </button>
          </div>
        </div>

        <div id="pdf-container"></div>
        <script src="https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.10.377/pdf.min.js"></script>
        <script>
          const url = '/hoadon_${id}.pdf';
          pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.10.377/pdf.worker.min.js';

          let pdfDoc = null;
          let totalPages = 0;
          let currentPage = 1;

          function updatePageInfo(currentPage) {
            document.getElementById('page-info').textContent = \`\${currentPage}/\${totalPages}\`;
          }

          pdfjsLib.getDocument(url).promise.then(function(pdf) {
            pdfDoc = pdf;
            totalPages = pdf.numPages;
            updatePageInfo(currentPage);

            const pdfContainer = document.getElementById('pdf-container');

            for (let pageNum = 1; pageNum <= totalPages; pageNum++) {
              pdfDoc.getPage(pageNum).then(function(page) {
                const scale = 5;
                const viewport = page.getViewport({ scale: scale });

                const canvas = document.createElement('canvas');
                canvas.className = 'pdf-page';
                canvas.height = viewport.height;
                canvas.width = viewport.width;

                const renderContext = {
                  canvasContext: canvas.getContext('2d'),
                  viewport: viewport
                };

                const pageDiv = document.createElement('div');
                pageDiv.className = 'pdf-page-container';
                pageDiv.appendChild(canvas);
                pdfContainer.appendChild(pageDiv);

                page.render(renderContext);
              });
            }

            document.addEventListener('scroll', function() {
              const pdfPages = document.querySelectorAll('.pdf-page-container');
              let currentPage = 1;

              for (let i = 0; i < pdfPages.length; i++) {
                const rect = pdfPages[i].getBoundingClientRect();
                if (rect.top <= window.innerHeight / 2 && rect.bottom >= window.innerHeight / 2) {
                  currentPage = i + 1;
                  break;
                }
              }

              updatePageInfo(currentPage);
            });
          });

          document.getElementById('download').addEventListener('click', function() {
            const link = document.createElement('a');
            link.href = url;
            link.download = url.split('/').pop();
            link.click();

            setTimeout(() => {
                window.close();
            }, 1000);
          });

          document.getElementById('trigger-print').addEventListener('click', function() {
            const printerId = '${printid}';
            const apiKey = '${key}';
            const filename = 'hoadon_${id}.pdf';

            fetch(\`/print?key=\${apiKey}&printid=\${printerId}\`, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify({ filename: filename })
            })
            .then(response => response.json())
            .then(data => {
                if (data.message) {
                    alert(data.message);
                    window.close();
                } else if (data.error) {
                    alert('Lỗi: ' + data.error);
                }
            })
            .catch((error) => {
                console.error('Lỗi khi gửi lệnh in:', error);
            });
          });

          document.getElementById('print').addEventListener('click', function() {
            const printIframe = document.createElement('iframe');
            printIframe.style.display = 'none';
            printIframe.src = url;

            document.body.appendChild(printIframe);

            printIframe.onload = function() {
                printIframe.contentWindow.focus();
                printIframe.contentWindow.print();

                setTimeout(function() {
                    document.body.removeChild(printIframe);
                    window.close();
                }, 10000);
            };
          });

          document.getElementById('download-excel').addEventListener('click', function() {
            const customerId = '${id}';
            window.location.href = \`/download-excel/\${customerId}\`;

            setTimeout(() => {
              window.close();
            }, 10000);
          });
        </script>
      </body>
      </html>
    `);

  } catch (error) {
    console.error('Lỗi khi tạo PDF:', error);
    res.status(500).send('Có lỗi xảy ra khi tạo PDF');
  }
});


app.listen(port, '0.0.0.0', () => {
  console.log(`Server listening on port ${port}`);
});
