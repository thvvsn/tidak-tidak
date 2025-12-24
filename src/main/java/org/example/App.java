package org.example; // Package tempat file ini berada

//library Apache POI untuk membuat file Excel
import org.apache.poi.ss.usermodel.*;          // Untuk objek Excel (Workbook, Sheet, Row)
import org.apache.poi.xssf.usermodel.XSSFWorkbook; // Untuk membuat file Excel .xlsx

//library Java
import java.io.File;                // Representasi file/folder
import java.io.FileOutputStream;    // Untuk menulis ke file
import java.io.IOException;         // Menangani error file
import java.util.*;                 // Struktur data (List, Set, Scanner, dll)


public class App {

    public static void main(String[] args) {
        Scanner input = new Scanner(System.in);

        // Set berfungsi menampung nama agar tidak terjadi duplikasi
        Set<String> daftarNama = new HashSet<>();

        // List untuk menyimpan seluruh data mahasiswa dalam bentuk array String
        List<String[]> dataMahasiswa = new ArrayList<>();

        System.out.println("Masukkan data mahasiswa. Ketik 'selesai' pada nama untuk mengakhiri");

        while (true) {
            System.out.print("\nMasukkan Nama: ");
            String nama = input.nextLine().trim();

            if (nama.equalsIgnoreCase("selesai")) {
                System.out.println("Terima kasih !");
                break;
            }

            // Cek apakah nama sudah pernah dimasukkan
            if (daftarNama.contains(nama)) {
                System.out.println("Nama sudah ada, masukkan nama yang berbeda !");
                continue;
            }

            // Input semester
            System.out.print("Masukkan Semester: ");
            String semester = input.nextLine().trim();

            // Input mata kuliah
            System.out.print("Masukkan Mata Kuliah: ");
            String matkul = input.nextLine().trim();

            // Menambahkan nama ke dalam Set (agar tidak duplikat)
            daftarNama.add(nama);

            // Menyimpan data mahasiswa ke dalam list
            dataMahasiswa.add(new String[]{nama, semester, matkul});

            // Menyimpan semua data ke file Excel setiap ada data baru
            simpanKeExcel(dataMahasiswa);

            System.out.println("Data berhasil disimpan ke dalam file data_mahasiswa.xlsx !");
        }
    }

    // Fungsi untuk menyimpan seluruh data mahasiswa ke file Excel
    public static void simpanKeExcel(List<String[]> data) {

        // Membuat workbook Excel baru (format .xlsx)
        Workbook workbook = new XSSFWorkbook();

        // Membuat sheet baru bernama "Data Mahasiswa"
        Sheet sheet = workbook.createSheet("Data Mahasiswa");

        // Membuat baris pertama sebagai header
        Row header = sheet.createRow(0);
        header.createCell(0).setCellValue("Nama");        // Kolom 1
        header.createCell(1).setCellValue("Semester");    // Kolom 2
        header.createCell(2).setCellValue("Mata Kuliah"); // Kolom 3

        // Mengisi data mulai baris kedua (index 1)
        int rowNum = 1; // Row index
        for (String[] rowData : data) {
            Row row = sheet.createRow(rowNum++); // Membuat baris baru

            row.createCell(0).setCellValue(rowData[0]); // Isi nama
            row.createCell(1).setCellValue(rowData[1]); // Isi semester
            row.createCell(2).setCellValue(rowData[2]); // Isi mata kuliah
        }

        try {
            // Menentukan folder penyimpanan yaitu src/main/resources
            File folder = new File("src/main/resources");

            // Jika folder belum ada, maka dibuat otomatis
            if (!folder.exists()) {
                folder.mkdirs();
            }

            // Membuat file output Excel dengan nama "data_mahasiswa.xlsx"
            FileOutputStream fileOut = new FileOutputStream("src/main/resources/data_mahasiswa.xlsx");

            // Menulis workbook ke dalam file output
            workbook.write(fileOut);

            // Menutup file dan workbook agar tidak terjadi memory leak
            fileOut.close();
            workbook.close();

        } catch (IOException e) {
            // Menangani error jika terjadi masalah saat menyimpan file
            System.out.println("Terjadi kesalahan saat menyimpan file: " + e.getMessage());
        }
    }
}
