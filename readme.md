# Aplikasi Analisis Daftar Pustaka

Halo!, Saya Nadir Mubarak Mahasiswa S1 Sistem Informasi Universitas Airlangga, angkatan 2018. Berikut ini adalah aplikasi hasil penelitian skripsi Saya yang diberi nama "Aplikasi Analisis Daftar Pustaka". Aplikasi ini dikembangkan berbasis website dengan bahasa pemrograman python dan framework Django. Jika anda ingin menggunakan aplikasi ini pastikan telah menginstall python sebelumnya.

# Data Data Sumber

Dalam berjalannya fungsi aplikasi, digunakan banyak data, seperti data jurnal terindeks scopus dan lainnya. Semua data tersebut tertulis lengkap di [https://docs.google.com/spreadsheets/d/11rCV7Q2WoKE1q88xDlX5a7B4Lx1kBM4Iiaj6TtPgXjw/edit?usp=sharing](https://docs.google.com/spreadsheets/d/11rCV7Q2WoKE1q88xDlX5a7B4Lx1kBM4Iiaj6TtPgXjw/edit?usp=sharing)

# Instalasi

 1. Masuk atau buat direktori untuk menginstall aplikasi
	 ```sh
	mkdir *nama_direktori/ cd *direktori
	```
2. Clone Repository
	 ```sh
	git clone https://github.com/nadirelc/skripsi.git
	```
3. Masuk kedalam Folder Repository
	 ```sh
	cd skripsi
	```
4. Install library/package yang diperlukan
	 ```sh
	pip install -r requirements.txt
	```
5. Jalankan aplikasi
	 ```sh
	python manage.py runserver
	```
6. Akses Website melalui localhost
	```
	Akses aplikasi web di browser http://127.0.0.1:8000/
	```

# Petunjuk Penggunaan
1. Masukkan Nama Author yang ada pada paper yang ingin dianalisis
2. Masukkan file paper berformat pdf
3. Klik submit, lalu file hasil analisis berbentuk excel akan otomatis terunduh

All your files and folders are presented as a tree in the file explorer. You can switch from one to another by clicking a file in the tree.

# Informasi peneliti
<a  href="https://www.linkedin.com/in/nadirelc/"  target="_blank"><img  src="https://user-images.githubusercontent.com/67138576/121289494-34a73d80-c90f-11eb-8811-7904e7b88606.png"  width="90"  height="90"></a> 
Nadir Mubarak
Instagarm : @nadirelc
Linkedin : https://www.linkedin.com/in/nadirelc/