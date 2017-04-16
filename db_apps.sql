/*
SQLyog Enterprise - MySQL GUI v8.05 
MySQL - 5.5.5-10.1.21-MariaDB : Database - db_apps
*********************************************************************
*/

/*!40101 SET NAMES utf8 */;

/*!40101 SET SQL_MODE=''*/;

/*!40014 SET @OLD_FOREIGN_KEY_CHECKS=@@FOREIGN_KEY_CHECKS, FOREIGN_KEY_CHECKS=0 */;
/*!40101 SET @OLD_SQL_MODE=@@SQL_MODE, SQL_MODE='NO_AUTO_VALUE_ON_ZERO' */;

CREATE DATABASE /*!32312 IF NOT EXISTS*/`db_apps` /*!40100 DEFAULT CHARACTER SET latin1 */;

USE `db_apps`;

/*Table structure for table `tbl_biaya_paket` */

DROP TABLE IF EXISTS `tbl_biaya_paket`;

CREATE TABLE `tbl_biaya_paket` (
  `id_biaya_paket` int(5) NOT NULL AUTO_INCREMENT,
  `kode_paket` varchar(10) DEFAULT NULL,
  `nama_paket` varchar(10) DEFAULT NULL,
  `paket_pertemuan` int(3) DEFAULT NULL,
  `biaya_paket` int(10) DEFAULT NULL,
  `biaya_daftar` int(10) DEFAULT NULL,
  `biaya_sim` int(10) DEFAULT NULL,
  PRIMARY KEY (`id_biaya_paket`)
) ENGINE=InnoDB AUTO_INCREMENT=7 DEFAULT CHARSET=latin1;

/*Data for the table `tbl_biaya_paket` */

insert  into `tbl_biaya_paket`(`id_biaya_paket`,`kode_paket`,`nama_paket`,`paket_pertemuan`,`biaya_paket`,`biaya_daftar`,`biaya_sim`) values (5,'PSC003','ADVANCE',5,280000,2000000,8000000),(6,'PSC004','BEGINNER',10,350000,3000000,8000000);

/*Table structure for table `tbl_history` */

DROP TABLE IF EXISTS `tbl_history`;

CREATE TABLE `tbl_history` (
  `id_history` int(5) NOT NULL AUTO_INCREMENT,
  `id_user` int(5) DEFAULT NULL,
  `id_level` int(1) DEFAULT NULL,
  `jam` time DEFAULT NULL,
  `tanggal` date DEFAULT NULL,
  PRIMARY KEY (`id_history`)
) ENGINE=InnoDB AUTO_INCREMENT=85 DEFAULT CHARSET=latin1;

/*Data for the table `tbl_history` */

insert  into `tbl_history`(`id_history`,`id_user`,`id_level`,`jam`,`tanggal`) values (1,1,1,'01:38:21','2017-03-28'),(2,2,2,'01:48:45','2017-03-28'),(3,1,1,'01:55:13','2017-03-28'),(4,1,1,'01:58:01','2017-03-28'),(5,1,1,'01:58:52','2017-03-28'),(6,1,1,'02:00:19','2017-03-28'),(7,1,1,'02:05:48','2017-03-28'),(8,1,1,'02:07:58','2017-03-28'),(9,1,1,'02:08:58','2017-03-28'),(10,1,1,'02:10:32','2017-03-28'),(11,1,1,'02:12:07','2017-03-28'),(12,1,1,'02:22:43','2017-03-28'),(13,2,2,'02:31:05','2017-03-28'),(14,2,2,'02:31:57','2017-03-28'),(15,2,2,'02:32:42','2017-03-28'),(16,2,2,'02:33:10','2017-03-28'),(17,1,1,'02:33:24','2017-03-28'),(18,1,1,'02:34:27','2017-03-28'),(19,1,1,'02:34:51','2017-03-28'),(20,2,2,'02:35:05','2017-03-28'),(21,1,1,'02:35:19','2017-03-28'),(22,2,2,'02:35:26','2017-03-28'),(23,1,1,'02:36:07','2017-03-28'),(24,2,2,'02:36:15','2017-03-28'),(25,1,1,'03:26:08','2017-03-28'),(26,1,1,'03:31:47','2017-03-28'),(27,1,1,'03:33:05','2017-03-28'),(28,1,1,'03:34:22','2017-03-28'),(29,1,1,'03:35:37','2017-03-28'),(30,1,1,'03:36:28','2017-03-28'),(31,1,1,'03:37:02','2017-03-28'),(32,1,1,'03:37:16','2017-03-28'),(33,1,1,'14:00:51','2017-03-28'),(34,1,1,'14:08:40','2017-03-28'),(35,1,1,'14:10:08','2017-03-28'),(36,1,1,'14:48:42','2017-03-28'),(37,1,1,'14:50:05','2017-03-28'),(38,1,1,'14:52:45','2017-03-28'),(39,1,1,'14:56:45','2017-03-28'),(40,1,1,'15:00:18','2017-03-28'),(41,1,1,'15:00:55','2017-03-28'),(42,1,1,'15:01:21','2017-03-28'),(43,1,1,'15:03:36','2017-03-28'),(44,1,1,'15:04:06','2017-03-28'),(45,1,1,'15:05:31','2017-03-28'),(46,1,1,'15:06:19','2017-03-28'),(47,1,1,'15:07:44','2017-03-28'),(48,1,1,'15:09:40','2017-03-28'),(49,1,1,'15:13:26','2017-03-28'),(50,18,1,'19:37:40','2017-03-28'),(51,18,1,'19:39:06','2017-03-28'),(52,1,1,'19:42:33','2017-03-28'),(53,1,1,'22:12:02','2017-03-29'),(54,1,1,'22:13:51','2017-03-29'),(55,1,1,'02:38:00','2017-03-30'),(56,1,1,'02:51:28','2017-03-30'),(57,1,1,'01:57:47','2017-04-01'),(58,1,1,'01:58:10','2017-04-01'),(59,1,1,'02:03:16','2017-04-01'),(60,1,1,'02:04:06','2017-04-01'),(61,1,1,'02:04:54','2017-04-01'),(62,1,1,'02:05:22','2017-04-01'),(63,1,1,'15:59:14','2017-04-02'),(64,1,1,'18:18:45','2017-04-03'),(65,1,1,'18:19:39','2017-04-03'),(66,1,1,'18:03:27','2017-04-05'),(67,1,1,'18:04:39','2017-04-05'),(68,1,1,'18:17:24','2017-04-10'),(69,1,1,'18:19:20','2017-04-10'),(70,1,1,'18:20:25','2017-04-10'),(71,1,1,'01:36:47','2017-04-16'),(72,1,1,'02:21:20','2017-04-16'),(73,1,1,'02:31:34','2017-04-16'),(74,1,1,'02:45:37','2017-04-16'),(75,1,1,'13:09:26','2017-04-16'),(76,1,1,'13:27:46','2017-04-16'),(77,1,1,'16:24:08','2017-04-16'),(78,1,1,'16:25:12','2017-04-16'),(79,1,1,'16:25:41','2017-04-16'),(80,1,1,'16:42:55','2017-04-16'),(81,2,2,'16:45:03','2017-04-16'),(82,2,2,'16:46:52','2017-04-16'),(83,2,2,'16:47:30','2017-04-16'),(84,1,1,'17:11:53','2017-04-16');

/*Table structure for table `tbl_jadwal` */

DROP TABLE IF EXISTS `tbl_jadwal`;

CREATE TABLE `tbl_jadwal` (
  `id_jadwal` int(5) NOT NULL AUTO_INCREMENT,
  `kode_jadwal` varchar(20) DEFAULT NULL,
  `no_registrasi` varchar(20) DEFAULT NULL,
  `kode_jam` varchar(20) DEFAULT NULL,
  `kode_mobil` varchar(20) DEFAULT NULL,
  PRIMARY KEY (`id_jadwal`)
) ENGINE=InnoDB AUTO_INCREMENT=5 DEFAULT CHARSET=latin1;

/*Data for the table `tbl_jadwal` */

insert  into `tbl_jadwal`(`id_jadwal`,`kode_jadwal`,`no_registrasi`,`kode_jam`,`kode_mobil`) values (3,'JSC001','TRSC1704002','JLSC005','SC001'),(4,'JSC002','TRSC1704001','JLSC005','SC001');

/*Table structure for table `tbl_jam_latihan` */

DROP TABLE IF EXISTS `tbl_jam_latihan`;

CREATE TABLE `tbl_jam_latihan` (
  `id_jam_latihan` int(5) NOT NULL AUTO_INCREMENT,
  `kode_jam_latihan` varchar(10) DEFAULT NULL,
  `hari` varchar(7) DEFAULT NULL,
  `durasi` int(11) DEFAULT NULL,
  PRIMARY KEY (`id_jam_latihan`)
) ENGINE=InnoDB AUTO_INCREMENT=6 DEFAULT CHARSET=latin1;

/*Data for the table `tbl_jam_latihan` */

insert  into `tbl_jam_latihan`(`id_jam_latihan`,`kode_jam_latihan`,`hari`,`durasi`) values (5,'JLSC005','JUMAT',6);

/*Table structure for table `tbl_mobil` */

DROP TABLE IF EXISTS `tbl_mobil`;

CREATE TABLE `tbl_mobil` (
  `id_mobil` int(5) NOT NULL AUTO_INCREMENT,
  `kode_mobil` varchar(10) DEFAULT NULL,
  `merek_mobil` varchar(20) DEFAULT NULL,
  `tipe_mobil` varchar(20) DEFAULT NULL,
  `plat_mobil` varchar(8) DEFAULT NULL,
  PRIMARY KEY (`id_mobil`)
) ENGINE=InnoDB AUTO_INCREMENT=2 DEFAULT CHARSET=latin1;

/*Data for the table `tbl_mobil` */

insert  into `tbl_mobil`(`id_mobil`,`kode_mobil`,`merek_mobil`,`tipe_mobil`,`plat_mobil`) values (1,'SC001','HONDA','JAZZ','B2054TWS');

/*Table structure for table `tbl_registrasi` */

DROP TABLE IF EXISTS `tbl_registrasi`;

CREATE TABLE `tbl_registrasi` (
  `id_registrasi` int(5) NOT NULL AUTO_INCREMENT,
  `no_registrasi` varchar(20) DEFAULT NULL,
  `kelas` varchar(10) DEFAULT NULL,
  `kode_paket` int(5) DEFAULT NULL,
  `total_bayar` varchar(20) DEFAULT NULL,
  `noinduk_siswa` varchar(20) DEFAULT NULL,
  `nama_siswa` varchar(20) DEFAULT NULL,
  `ktp` varchar(20) DEFAULT NULL,
  `telepon` varchar(20) DEFAULT NULL,
  `email` varchar(20) DEFAULT NULL,
  `alamat` text,
  `flag_jadwal` int(1) NOT NULL DEFAULT '1',
  PRIMARY KEY (`id_registrasi`)
) ENGINE=InnoDB AUTO_INCREMENT=3 DEFAULT CHARSET=latin1;

/*Data for the table `tbl_registrasi` */

insert  into `tbl_registrasi`(`id_registrasi`,`no_registrasi`,`kelas`,`kode_paket`,`total_bayar`,`noinduk_siswa`,`nama_siswa`,`ktp`,`telepon`,`email`,`alamat`,`flag_jadwal`) values (1,'TRSC1704001','PAGI',5,'10280000','NSC001','POPO','1527380181871862','081278870852','me@suwondo.id','asdadsads',1),(2,'TRSC1704002','SIANG',6,'11350000','NSC002','IPUNK','1289128712644912','9889398893','mail@suwondo.id','askjdasdasd',1);

/*Table structure for table `tbl_siswa` */

DROP TABLE IF EXISTS `tbl_siswa`;

CREATE TABLE `tbl_siswa` (
  `id_siswa` int(5) NOT NULL AUTO_INCREMENT,
  `noinduk_siswa` varchar(10) DEFAULT NULL,
  `nama_siswa` varchar(20) DEFAULT NULL,
  `alamat_siswa` text,
  `ktp_siswa` varchar(20) DEFAULT NULL,
  `telpon_siswa` varchar(20) DEFAULT NULL,
  `email_siswa` varchar(30) DEFAULT NULL,
  PRIMARY KEY (`id_siswa`)
) ENGINE=InnoDB AUTO_INCREMENT=11 DEFAULT CHARSET=latin1;

/*Data for the table `tbl_siswa` */

insert  into `tbl_siswa`(`id_siswa`,`noinduk_siswa`,`nama_siswa`,`alamat_siswa`,`ktp_siswa`,`telpon_siswa`,`email_siswa`) values (7,'NSC001','POPO','jakarta','1231312312321132','081278870852','mail.com'),(8,'NSC002','IPUNK','rawa bebek','1231312312321132','081278870852','mail.com'),(9,'NSC003','popo','asdadsadad','123123213','2123123','sdasdass'),(10,'NSC004','adad','daadsad','121212','212131','asdsd');

/*Table structure for table `tbl_user` */

DROP TABLE IF EXISTS `tbl_user`;

CREATE TABLE `tbl_user` (
  `id_user` int(5) NOT NULL AUTO_INCREMENT,
  `namauser` varchar(20) DEFAULT NULL,
  `username` varchar(10) DEFAULT NULL,
  `katasandi` varchar(10) DEFAULT NULL,
  `level` int(1) DEFAULT NULL,
  PRIMARY KEY (`id_user`)
) ENGINE=InnoDB AUTO_INCREMENT=3 DEFAULT CHARSET=latin1;

/*Data for the table `tbl_user` */

insert  into `tbl_user`(`id_user`,`namauser`,`username`,`katasandi`,`level`) values (1,'Administrator','admin','admin',1),(2,'User biasa','user','user',2);

/*!40101 SET SQL_MODE=@OLD_SQL_MODE */;
/*!40014 SET FOREIGN_KEY_CHECKS=@OLD_FOREIGN_KEY_CHECKS */;
