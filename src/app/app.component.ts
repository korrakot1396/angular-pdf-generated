import { Component, ViewChild, ElementRef } from '@angular/core';
import jsPDF from 'jspdf';
import html2pdf from 'html2pdf.js';
import * as XLSX from 'xlsx';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css'],
})
export class AppComponent {
  @ViewChild('htmlData') htmlData!: ElementRef;

  constructor() {}

  // ฟังก์ชันนี้เปิดเอกสาร PDF เมื่อถูกเรียกใช้
  public openPDF(): void {
    // รับองค์ประกอบ HTML ที่มี ID 'htmlData'
    const HTML_CONTENT = document.getElementById('htmlData');
    // สร้างอินสแตนซ์ของ jsPDF
    const pdf = new jsPDF();
    pdf.setFontSize(12);
    // ใช้ไลบรารี html2pdf เพื่อแปลงเนื้อหา HTML เป็น PDF
    html2pdf()
      .from(HTML_CONTENT)
      .set({
        filename: 'ใบแสดงรายการเบิก.pdf', // กำหนดชื่อไฟล์ของ PDF
        margin: [10, 10], // กำหนดระยะขอบของ PDF
        jsPDF: pdf, // กำหนดอินสแตนซ์ jsPDF
      })
      .toPdf()
      .get('pdf')
      .then(function (pdfDoc: jsPDF | any) {
        const totalPages = pdfDoc.internal.getNumberOfPages();

        // วนลูปผ่านหน้าทั้งหมดของเอกสาร PDF
        for (let i = totalPages; i >= 1; i--) {
          pdfDoc.setPage(i);

          const pageContent = pdfDoc.output('dataurlstring');

          // หากหน้านั้นไม่มีเนื้อหาใด ๆ ให้ลบหน้านั้นทิ้ง
          if (pageContent === 'data:,') {
            pdfDoc.deletePage(i);
          } else {
            break;
          }
        }
        return pdfDoc;
      })
      .save(); // บันทึกเอกสาร PDF
  }

  // ฟังก์ชันนี้ส่งออกข้อมูลในรูปแบบไฟล์ Excel เมื่อถูกเรียกใช้
  public exportToExcel(): void {
    // รับองค์ประกอบตาราง HTML ที่มี ID 'htmlData'
    const table = document.getElementById('htmlData') as HTMLTableElement;

    // แปลงตารางเป็น worksheet
    const worksheet: XLSX.WorkSheet = XLSX.utils.table_to_sheet(table);

    // สร้าง workbook ที่มีชื่อเป็น 'data' และมีข้อมูล worksheet
    const workbook: XLSX.WorkBook = { Sheets: { 'data': worksheet }, SheetNames: ['data'] };

    // เขียนข้อมูลเป็นไฟล์ Excel และเก็บไว้ใน excelBuffer
    const excelBuffer: any = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });

    // บันทึกไฟล์ Excel
    this.saveAsExcelFile(excelBuffer, 'ใบแสดงรายการเบิก');
  }

  // ฟังก์ชันเอาไว้สำหรับบันทึกไฟล์ Excel
  private saveAsExcelFile(buffer: any, fileName: string): void {
    // สร้าง Blob จากข้อมูลในรูปแบบ buffer
    const data: Blob = new Blob([buffer], { type: 'application/octet-stream' });

    // สร้าง URL สำหรับ Blob ที่สร้างขึ้น
    const url: string = window.URL.createObjectURL(data);

    // สร้างองค์ประกอบแบบ 'a' สำหรับลิงก์ดาวน์โหลดไฟล์
    const link: HTMLAnchorElement = document.createElement('a');
    link.href = url;
    link.download = `${fileName}.xlsx`; // กำหนดชื่อไฟล์ที่จะดาวน์โหลด
    link.click();

    // เพิ่มเวลาหลังจากคลิกลิงก์แล้ว จากนั้นเรียกใช้งาน window.URL.revokeObjectURL เพื่อล้าง URL
    setTimeout(() => {
      window.URL.revokeObjectURL(url);
    }, 100);
  }
}
