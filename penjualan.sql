-- MySQL dump 10.11
--
-- Host: localhost    Database: penjualan
-- ------------------------------------------------------
-- Server version	5.0.51b-community-nt-log

/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
/*!40101 SET NAMES utf8 */;
/*!40103 SET @OLD_TIME_ZONE=@@TIME_ZONE */;
/*!40103 SET TIME_ZONE='+00:00' */;
/*!40014 SET @OLD_UNIQUE_CHECKS=@@UNIQUE_CHECKS, UNIQUE_CHECKS=0 */;
/*!40014 SET @OLD_FOREIGN_KEY_CHECKS=@@FOREIGN_KEY_CHECKS, FOREIGN_KEY_CHECKS=0 */;
/*!40101 SET @OLD_SQL_MODE=@@SQL_MODE, SQL_MODE='NO_AUTO_VALUE_ON_ZERO' */;
/*!40111 SET @OLD_SQL_NOTES=@@SQL_NOTES, SQL_NOTES=0 */;

--
-- Table structure for table `akun`
--
Drop database if exists penjualan;
create database penjualan;
DROP TABLE IF EXISTS `akun`;
SET @saved_cs_client     = @@character_set_client;
SET character_set_client = utf8;
CREATE TABLE `akun` (
  `no_akun` varchar(10) NOT NULL default '',
  `nama_akun` varchar(255) default NULL,
  `jns` varchar(255) default NULL,
  PRIMARY KEY  (`no_akun`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1 ROW_FORMAT=DYNAMIC;
SET character_set_client = @saved_cs_client;

--
-- Dumping data for table `akun`
--

LOCK TABLES `akun` WRITE;
/*!40000 ALTER TABLE `akun` DISABLE KEYS */;
INSERT INTO `akun` VALUES ('001','Modal','2'),('100','Kas','1'),('101','Utang Dagang','2'),('131','Piutang dagang','1'),('200','Retur Jual','2'),('300','Pembelian barang','1'),('401','Beban-beban','1'),('480','Iklan','1'),('700','Penjualan','2'),('702','Pemasukan Lain-lain','2');
/*!40000 ALTER TABLE `akun` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `bagian`
--

DROP TABLE IF EXISTS `bagian`;
SET @saved_cs_client     = @@character_set_client;
SET character_set_client = utf8;
CREATE TABLE `bagian` (
  `Id` int(11) NOT NULL auto_increment,
  PRIMARY KEY  (`Id`)
) ENGINE=MyISAM DEFAULT CHARSET=utf8 ROW_FORMAT=FIXED;
SET character_set_client = @saved_cs_client;

--
-- Dumping data for table `bagian`
--

LOCK TABLES `bagian` WRITE;
/*!40000 ALTER TABLE `bagian` DISABLE KEYS */;
/*!40000 ALTER TABLE `bagian` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `byr_hutang`
--

DROP TABLE IF EXISTS `byr_hutang`;
SET @saved_cs_client     = @@character_set_client;
SET character_set_client = utf8;
CREATE TABLE `byr_hutang` (
  `Tanggal` date default NULL,
  `Jumlah_byr` double(19,4) default NULL,
  `No_pembelian` varchar(20) default NULL,
  `Id_supplier` varchar(20) default NULL,
  `Kode_bank` varchar(50) default NULL,
  `Bentuk` varchar(50) default NULL,
  `No_giro` varchar(50) default NULL,
  `no_bayar` varchar(50) default NULL,
  KEY `Id_supplier` (`Id_supplier`),
  KEY `Pembelianbyr_hutang` (`No_pembelian`)
) ENGINE=MyISAM DEFAULT CHARSET=utf8 ROW_FORMAT=DYNAMIC;
SET character_set_client = @saved_cs_client;

--
-- Dumping data for table `byr_hutang`
--

LOCK TABLES `byr_hutang` WRITE;
/*!40000 ALTER TABLE `byr_hutang` DISABLE KEYS */;
/*!40000 ALTER TABLE `byr_hutang` ENABLE KEYS */;
UNLOCK TABLES;

/*!50003 SET @SAVE_SQL_MODE=@@SQL_MODE*/;

DELIMITER ;;
/*!50003 SET SESSION SQL_MODE="NO_AUTO_CREATE_USER,NO_ENGINE_SUBSTITUTION" */;;
/*!50003 CREATE */ /*!50017 DEFINER=`root`@`localhost` */ /*!50003 TRIGGER `tbyrhutangai` AFTER INSERT ON `byr_hutang` FOR EACH ROW BEGIN
update hutang set jumlah_byr=jumlah_byr+new.jumlah_byr where no_pembelian=new.no_pembelian;
insert into keuangan(tanggal,keterangan,pengeluaran,jenis,no_transaksi) values(new.tanggal,'Pembayaran hutang',new.jumlah_byr,'Pembayaran hutang',new.no_bayar);
update tblsupplier set jumlah_hutang=jumlah_hutang-new.jumlah_byr where id_supplier=new.id_supplier;
END */;;

/*!50003 SET SESSION SQL_MODE="NO_AUTO_CREATE_USER,NO_ENGINE_SUBSTITUTION" */;;
/*!50003 CREATE */ /*!50017 DEFINER=`root`@`localhost` */ /*!50003 TRIGGER `tbyrhutang` AFTER DELETE ON `byr_hutang` FOR EACH ROW BEGIN

update hutang set jumlah_byr=jumlah_byr-old.jumlah_byr where no_pembelian=old.no_pembelian;
delete from keuangan where no_transaksi=old.no_bayar;
update tblsupplier set jumlah_hutang=jumlah_hutang+old.jumlah_byr where id_supplier=old.id_supplier;
END */;;

DELIMITER ;
/*!50003 SET SESSION SQL_MODE=@SAVE_SQL_MODE*/;

--
-- Table structure for table `byr_piutang`
--

DROP TABLE IF EXISTS `byr_piutang`;
SET @saved_cs_client     = @@character_set_client;
SET character_set_client = utf8;
CREATE TABLE `byr_piutang` (
  `no_bayar` varchar(50) default NULL,
  `Tanggal` date default NULL,
  `Jumlah_byr` double(19,4) default NULL,
  `No_penjualan` varchar(20) default NULL,
  `Id_pelanggan` varchar(20) default NULL,
  `Kode_bank` varchar(50) default NULL,
  `Bentuk` varchar(50) default NULL,
  `No_giro` varchar(50) default NULL,
  KEY `Id_supplier` (`Id_pelanggan`),
  KEY `Penjualanbyr_piutang` (`No_penjualan`)
) ENGINE=MyISAM DEFAULT CHARSET=utf8 ROW_FORMAT=DYNAMIC;
SET character_set_client = @saved_cs_client;

--
-- Dumping data for table `byr_piutang`
--

LOCK TABLES `byr_piutang` WRITE;
/*!40000 ALTER TABLE `byr_piutang` DISABLE KEYS */;
/*!40000 ALTER TABLE `byr_piutang` ENABLE KEYS */;
UNLOCK TABLES;

/*!50003 SET @SAVE_SQL_MODE=@@SQL_MODE*/;

DELIMITER ;;
/*!50003 SET SESSION SQL_MODE="NO_AUTO_CREATE_USER,NO_ENGINE_SUBSTITUTION" */;;
/*!50003 CREATE */ /*!50017 DEFINER=`root`@`localhost` */ /*!50003 TRIGGER `tbyrpiutangai` AFTER INSERT ON `byr_piutang` FOR EACH ROW BEGIN
update piutang set jumlah_byr=jumlah_byr+new.jumlah_byr where no_penjualan=new.no_penjualan;
insert into keuangan(tanggal,keterangan,pemasukan,jenis,no_transaksi) values(new.tanggal,'Pembayaran piutang',new.jumlah_byr,'Pembayaran piutang',new.no_bayar);
update pelanggan set jumlah_piutang=jumlah_piutang-new.jumlah_byr where id_pelanggan=new.id_pelanggan;
END */;;

/*!50003 SET SESSION SQL_MODE="NO_AUTO_CREATE_USER,NO_ENGINE_SUBSTITUTION" */;;
/*!50003 CREATE */ /*!50017 DEFINER=`root`@`localhost` */ /*!50003 TRIGGER `tbyrpiutang` AFTER DELETE ON `byr_piutang` FOR EACH ROW BEGIN

update piutang set jumlah_byr=jumlah_byr-old.jumlah_byr where no_penjualan=old.no_penjualan;
delete from keuangan where no_transaksi=old.no_bayar;
update pelanggan set jumlah_piutang=jumlah_piutang+old.jumlah_byr where id_pelanggan=old.id_pelanggan;
END */;;

DELIMITER ;
/*!50003 SET SESSION SQL_MODE=@SAVE_SQL_MODE*/;

--
-- Table structure for table `data_toko`
--

DROP TABLE IF EXISTS `data_toko`;
SET @saved_cs_client     = @@character_set_client;
SET character_set_client = utf8;
CREATE TABLE `data_toko` (
  `Nama_toko` varchar(60) NOT NULL default '',
  `Alamat` varchar(255) default NULL,
  `Kota` varchar(30) default NULL,
  `No_telp` varchar(50) default NULL,
  PRIMARY KEY  (`Nama_toko`)
) ENGINE=MyISAM DEFAULT CHARSET=utf8 ROW_FORMAT=DYNAMIC;
SET character_set_client = @saved_cs_client;

--
-- Dumping data for table `data_toko`
--

LOCK TABLES `data_toko` WRITE;
/*!40000 ALTER TABLE `data_toko` DISABLE KEYS */;
/*!40000 ALTER TABLE `data_toko` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `detilbeli`
--

DROP TABLE IF EXISTS `detilbeli`;
SET @saved_cs_client     = @@character_set_client;
SET character_set_client = utf8;
CREATE TABLE `detilbeli` (
  `No_Pembelian` varchar(20) default NULL,
  `kode_brg` varchar(20) default NULL,
  `Harga_beli` double(19,2) default NULL,
  `Jumlah_brg` double(10,2) default NULL,
  `diskon` double(19,2) default NULL,
  `Total` double(19,2) default NULL,
  `jumlah_brg2` double(10,2) default NULL,
  `satuan` varchar(50) default NULL,
  KEY `detilbeliKode_brg` (`kode_brg`),
  KEY `Pembeliandetilbeli` (`No_Pembelian`)
) ENGINE=MyISAM DEFAULT CHARSET=utf8 ROW_FORMAT=DYNAMIC;
SET character_set_client = @saved_cs_client;

--
-- Dumping data for table `detilbeli`
--

LOCK TABLES `detilbeli` WRITE;
/*!40000 ALTER TABLE `detilbeli` DISABLE KEYS */;
/*!40000 ALTER TABLE `detilbeli` ENABLE KEYS */;
UNLOCK TABLES;

/*!50003 SET @SAVE_SQL_MODE=@@SQL_MODE*/;

DELIMITER ;;
/*!50003 SET SESSION SQL_MODE="NO_AUTO_CREATE_USER,NO_ENGINE_SUBSTITUTION" */;;
/*!50003 CREATE */ /*!50017 DEFINER=`root`@`localhost` */ /*!50003 TRIGGER `tdetilbeli` AFTER INSERT ON `detilbeli` FOR EACH ROW BEGIN
declare ttl double;
declare disk double;
set ttl=(select coalesce(sum(`Harga_beli` * `jumlah_brg`),0)from detilbeli where `No_pembelian`=new.no_pembelian);
set disk=(select coalesce(sum(diskon),0)from detilbeli where `No_Pembelian`=new.no_pembelian);
set @ttl=ttl;set @disk=disk;
update pembelian set `total`=@ttl,total_diskon=total_diskon+@disk where `no_pembelian`=new.no_pembelian;
update tblbarang set stok=stok+new.jumlah_brg where kode_brg=new.kode_brg;
END */;;

/*!50003 SET SESSION SQL_MODE="NO_AUTO_CREATE_USER,NO_ENGINE_SUBSTITUTION" */;;
/*!50003 CREATE */ /*!50017 DEFINER=`root`@`localhost` */ /*!50003 TRIGGER `tdetilbeli2` AFTER DELETE ON `detilbeli` FOR EACH ROW BEGIN
declare ttl double;declare cek double (19,2);
declare disk double;
set ttl=(select coalesce(sum(`Harga_beli` * `jumlah_brg`),0)from detilbeli where `No_Pembelian`=old.No_Pembelian);
set disk=(select coalesce(sum(diskon),0)from detilbeli where `No_Pembelian`=old.No_Pembelian);
set @ttl=ttl;set @disk=disk;

set cek=(select coalesce(total,0) from pembelian where no_pembelian=old.No_Pembelian);
if cek!=0 then
update pembelian set `total`=@ttl,total_diskon=total_diskon+@disk where `No_pembelian`=old.No_Pembelian;
end if;
update tblbarang set stok=stok-old.jumlah_brg where kode_brg=old.kode_brg;
END */;;

DELIMITER ;
/*!50003 SET SESSION SQL_MODE=@SAVE_SQL_MODE*/;

--
-- Table structure for table `detiljual`
--

DROP TABLE IF EXISTS `detiljual`;
SET @saved_cs_client     = @@character_set_client;
SET character_set_client = utf8;
CREATE TABLE `detiljual` (
  `No_Penjualan` varchar(50) default NULL,
  `Kode_brg` varchar(20) default NULL,
  `Harga_jual` double(19,4) default NULL,
  `Harga_beli` double(19,4) default NULL,
  `Jumlah_brg` double(10,2) default NULL,
  `diskon` double(19,4) default NULL,
  `Total` double(19,4) default NULL,
  `jumlah_brg2` double(10,2) default NULL,
  `satuan` varchar(50) default NULL,
  KEY `detiljualKode_brg` (`Kode_brg`),
  KEY `detiljualNo_penjualan` (`No_Penjualan`),
  KEY `Penjualandetiljual` (`No_Penjualan`)
) ENGINE=MyISAM DEFAULT CHARSET=utf8 ROW_FORMAT=DYNAMIC;
SET character_set_client = @saved_cs_client;

--
-- Dumping data for table `detiljual`
--

LOCK TABLES `detiljual` WRITE;
/*!40000 ALTER TABLE `detiljual` DISABLE KEYS */;
/*!40000 ALTER TABLE `detiljual` ENABLE KEYS */;
UNLOCK TABLES;

/*!50003 SET @SAVE_SQL_MODE=@@SQL_MODE*/;

DELIMITER ;;
/*!50003 SET SESSION SQL_MODE="NO_AUTO_CREATE_USER,NO_ENGINE_SUBSTITUTION" */;;
/*!50003 CREATE */ /*!50017 DEFINER=`root`@`localhost` */ /*!50003 TRIGGER `tdetiljual` AFTER INSERT ON `detiljual` FOR EACH ROW BEGIN
declare ttl double;
declare disk double;
set ttl=(select coalesce(sum(`Harga_jual` * `jumlah_brg`),0)from detiljual where `No_Penjualan`=new.no_penjualan);
set disk=(select coalesce(sum(diskon),0)from detiljual where `No_Penjualan`=new.no_penjualan);
set @ttl=ttl;set @disk=disk;
update penjualan set `jumlah`=@ttl,total_diskon=total_diskon+@disk where `no_penjualan`=new.no_penjualan;
update tblbarang set stok=stok-new.jumlah_brg where kode_brg=new.kode_brg;
END */;;

/*!50003 SET SESSION SQL_MODE="NO_AUTO_CREATE_USER,NO_ENGINE_SUBSTITUTION" */;;
/*!50003 CREATE */ /*!50017 DEFINER=`root`@`localhost` */ /*!50003 TRIGGER `tdetiljual2` AFTER DELETE ON `detiljual` FOR EACH ROW BEGIN
declare ttl double;declare cek double (19,2);
declare disk double;
set ttl=(select coalesce(sum(`Harga_jual` * `jumlah_brg`),0)from detiljual where `No_Penjualan`=old.no_penjualan);
set disk=(select coalesce(sum(diskon),0)from detiljual where `No_Penjualan`=old.no_penjualan);
set @ttl=ttl;set @disk=disk;

set cek=(select coalesce(total,0) from penjualan where no_penjualan=old.no_penjualan);
if cek!=0 then
update penjualan set `jumlah`=@ttl,total_diskon=total_diskon+@disk where `no_penjualan`=old.no_penjualan;
end if;
update tblbarang set stok=stok+old.jumlah_brg where kode_brg=old.kode_brg;
END */;;

DELIMITER ;
/*!50003 SET SESSION SQL_MODE=@SAVE_SQL_MODE*/;

--
-- Table structure for table `detiljuals`
--

DROP TABLE IF EXISTS `detiljuals`;
SET @saved_cs_client     = @@character_set_client;
SET character_set_client = utf8;
CREATE TABLE `detiljuals` (
  `No_Penjualan` varchar(50) default NULL,
  `Kode_brg` varchar(20) default NULL,
  `Harga_jual` double(19,4) default NULL,
  `Harga_beli` double(19,4) default NULL,
  `Jumlah_brg` double(10,2) default NULL,
  `diskon` double(19,4) default NULL,
  `Total` double(19,4) default NULL,
  `Teretur` double(10,2) default NULL,
  `kembali_brg` double(10,2) default NULL,
  `kembali_uang` double(10,2) default NULL,
  `kembali_uang2` double(10,2) default NULL,
  `total_retur` double(19,4) default NULL,
  `jumlah_brg2` double(10,2) default NULL,
  `satuan` varchar(50) default NULL
) ENGINE=MyISAM DEFAULT CHARSET=utf8 ROW_FORMAT=DYNAMIC;
SET character_set_client = @saved_cs_client;

--
-- Dumping data for table `detiljuals`
--

LOCK TABLES `detiljuals` WRITE;
/*!40000 ALTER TABLE `detiljuals` DISABLE KEYS */;
/*!40000 ALTER TABLE `detiljuals` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `detilpesan`
--

DROP TABLE IF EXISTS `detilpesan`;
SET @saved_cs_client     = @@character_set_client;
SET character_set_client = utf8;
CREATE TABLE `detilpesan` (
  `No_pemesanan` varchar(20) NOT NULL default '',
  `Kode_brg` varchar(20) NOT NULL default '',
  `Jumlah_brg` double(10,2) default NULL,
  `jumlah_brg2` double(10,2) default NULL,
  `satuan` varchar(50) default NULL,
  `jumlah_beli` double(10,2) default NULL,
  PRIMARY KEY  (`No_pemesanan`,`Kode_brg`),
  KEY `detilpesanKode_brg` (`Kode_brg`),
  KEY `Pemesanandetilpesan` (`No_pemesanan`)
) ENGINE=MyISAM DEFAULT CHARSET=utf8 ROW_FORMAT=DYNAMIC;
SET character_set_client = @saved_cs_client;

--
-- Dumping data for table `detilpesan`
--

LOCK TABLES `detilpesan` WRITE;
/*!40000 ALTER TABLE `detilpesan` DISABLE KEYS */;
/*!40000 ALTER TABLE `detilpesan` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `detilreturbeli`
--

DROP TABLE IF EXISTS `detilreturbeli`;
SET @saved_cs_client     = @@character_set_client;
SET character_set_client = utf8;
CREATE TABLE `detilreturbeli` (
  `No_retur` varchar(20) NOT NULL default '',
  `No_pembelian` varchar(20) NOT NULL default '',
  `Kode_brg` varchar(20) NOT NULL default '',
  `Jumlah` double(10,2) default NULL,
  `Harga_beli` double(19,2) default NULL,
  `diskon` double(19,2) NOT NULL default '0.00',
  `Total` double(19,2) default NULL,
  `Alasan` varchar(255) default NULL,
  PRIMARY KEY  (`No_retur`,`No_pembelian`,`Kode_brg`),
  KEY `Retur_belidetil_returbeli` (`No_retur`)
) ENGINE=MyISAM DEFAULT CHARSET=utf8 ROW_FORMAT=DYNAMIC;
SET character_set_client = @saved_cs_client;

--
-- Dumping data for table `detilreturbeli`
--

LOCK TABLES `detilreturbeli` WRITE;
/*!40000 ALTER TABLE `detilreturbeli` DISABLE KEYS */;
/*!40000 ALTER TABLE `detilreturbeli` ENABLE KEYS */;
UNLOCK TABLES;

/*!50003 SET @SAVE_SQL_MODE=@@SQL_MODE*/;

DELIMITER ;;
/*!50003 SET SESSION SQL_MODE="NO_AUTO_CREATE_USER,NO_ENGINE_SUBSTITUTION" */;;
/*!50003 CREATE */ /*!50017 DEFINER=`root`@`localhost` */ /*!50003 TRIGGER `tdtlrtrbeli` BEFORE INSERT ON `detilreturbeli` FOR EACH ROW BEGIN
declare vharga double(19,2);declare vdisc double(19,2);declare vttl double(19,2);
set vharga=(Select harga_beli from detilbeli where no_pembelian=new.no_pembelian and kode_brg=new.kode_brg);
set vdisc=(select (diskon/jumlah_brg) from detilbeli where no_pembelian=new.no_pembelian and kode_brg=new.kode_brg);
set @vharga=vharga;set @vdisc=vdisc;
set new.harga_beli=@vharga;set new.diskon=@vdisc*new.jumlah;
set new.total=new.harga_beli*new.jumlah-new.diskon;
update tblbarang set stok=stok-new.jumlah where kode_brg=new.kode_brg;

END */;;

/*!50003 SET SESSION SQL_MODE="NO_AUTO_CREATE_USER,NO_ENGINE_SUBSTITUTION" */;;
/*!50003 CREATE */ /*!50017 DEFINER=`root`@`localhost` */ /*!50003 TRIGGER `tdtlrtrbeli3` AFTER INSERT ON `detilreturbeli` FOR EACH ROW BEGIN
declare vttl double(19,2);declare vttl2 double(19,2);
set vttl=(select coalesce(sum(`total`),0)from detilreturbeli where `No_retur`=new.No_retur);
set vttl2=(select coalesce(sum(`jumlah`),0)from detilreturbeli where `No_retur`=new.No_retur);
set @vttl=vttl;set @vttl2=vttl2;
update retur_beli set `total`=@vttl,total_brg=@vttl2 where `No_retur`=new.`No_retur`;
END */;;

/*!50003 SET SESSION SQL_MODE="NO_AUTO_CREATE_USER,NO_ENGINE_SUBSTITUTION" */;;
/*!50003 CREATE */ /*!50017 DEFINER=`root`@`localhost` */ /*!50003 TRIGGER `tdtlrtrbeli2` AFTER DELETE ON `detilreturbeli` FOR EACH ROW BEGIN
declare vttl double(19,2);
update tblbarang set stok=stok+old.jumlah where kode_brg=old.kode_brg;
set vttl=(select coalesce(sum(`total`),0) from detilreturbeli where `No_retur`=old.no_retur);
set @vttl=vttl;

END */;;

DELIMITER ;
/*!50003 SET SESSION SQL_MODE=@SAVE_SQL_MODE*/;

--
-- Table structure for table `detilreturjual`
--

DROP TABLE IF EXISTS `detilreturjual`;
SET @saved_cs_client     = @@character_set_client;
SET character_set_client = utf8;
CREATE TABLE `detilreturjual` (
  `No_retur` varchar(20) NOT NULL default '',
  `No_penjualan` varchar(50) NOT NULL default '',
  `Kode_brg` varchar(20) NOT NULL default '',
  `Jumlah` double(10,2) default NULL,
  `Harga_jual` double(19,4) default NULL,
  `diskon` double(19,2) NOT NULL default '0.00',
  `Total` double(19,4) default NULL,
  `Alasan` varchar(255) default NULL,
  PRIMARY KEY  (`No_retur`,`No_penjualan`,`Kode_brg`),
  KEY `retur_jualdetil_returjual` (`No_retur`)
) ENGINE=MyISAM DEFAULT CHARSET=utf8 ROW_FORMAT=DYNAMIC;
SET character_set_client = @saved_cs_client;

--
-- Dumping data for table `detilreturjual`
--

LOCK TABLES `detilreturjual` WRITE;
/*!40000 ALTER TABLE `detilreturjual` DISABLE KEYS */;
/*!40000 ALTER TABLE `detilreturjual` ENABLE KEYS */;
UNLOCK TABLES;

/*!50003 SET @SAVE_SQL_MODE=@@SQL_MODE*/;

DELIMITER ;;
/*!50003 SET SESSION SQL_MODE="NO_AUTO_CREATE_USER,NO_ENGINE_SUBSTITUTION" */;;
/*!50003 CREATE */ /*!50017 DEFINER=`root`@`localhost` */ /*!50003 TRIGGER `tdtlrtrjual` BEFORE INSERT ON `detilreturjual` FOR EACH ROW BEGIN
declare vharga double(19,2);declare vdisc double(19,2);declare vttl double(19,2);
set vharga=(Select harga_jual from detiljual where no_penjualan=new.no_penjualan and kode_brg=new.kode_brg);
set vdisc=(select (diskon/jumlah_brg) from detiljual where no_penjualan=new.no_penjualan and kode_brg=new.kode_brg);
set @vharga=vharga;set @vdisc=vdisc;
set new.harga_jual=@vharga;set new.diskon=@vdisc*new.jumlah;
set new.total=new.harga_jual*new.jumlah-new.diskon;
update tblbarang set stok=stok+new.jumlah where kode_brg=new.kode_brg;

END */;;

/*!50003 SET SESSION SQL_MODE="NO_AUTO_CREATE_USER,NO_ENGINE_SUBSTITUTION" */;;
/*!50003 CREATE */ /*!50017 DEFINER=`root`@`localhost` */ /*!50003 TRIGGER `tdtlrtrjual3` AFTER INSERT ON `detilreturjual` FOR EACH ROW BEGIN
declare vttl double(19,2);declare vttl2 double(19,2);
set vttl=(select coalesce(sum(`total`),0)from detilreturjual where `No_retur`=new.No_retur);
set vttl2=(select coalesce(sum(`jumlah`),0)from detilreturjual where `No_retur`=new.No_retur);
set @vttl=vttl;set @vttl2=vttl2;
update retur_jual set `total`=@vttl,total_brg=@vttl2 where `No_retur`=new.`No_retur`;
END */;;

/*!50003 SET SESSION SQL_MODE="NO_AUTO_CREATE_USER,NO_ENGINE_SUBSTITUTION" */;;
/*!50003 CREATE */ /*!50017 DEFINER=`root`@`localhost` */ /*!50003 TRIGGER `tdtlrtrjual2` AFTER DELETE ON `detilreturjual` FOR EACH ROW BEGIN
declare vttl double(19,2);
update tblbarang set stok=stok-old.jumlah where kode_brg=old.kode_brg;
set vttl=(select coalesce(sum(`total`),0) from detilreturjual where `No_retur`=old.no_retur);
set @vttl=vttl;

END */;;

DELIMITER ;
/*!50003 SET SESSION SQL_MODE=@SAVE_SQL_MODE*/;

--
-- Table structure for table `giro`
--

DROP TABLE IF EXISTS `giro`;
SET @saved_cs_client     = @@character_set_client;
SET character_set_client = utf8;
CREATE TABLE `giro` (
  `Tanggal` date default NULL,
  `No_giro` varchar(50) default NULL,
  `tgl_jt` date default NULL,
  `Kode_bank` varchar(50) default NULL,
  `giro_masuk` double(19,4) default NULL,
  `giro_keluar` double(19,4) default NULL,
  `No_faktur` varchar(50) default NULL,
  `Keterangan` varchar(255) default NULL,
  KEY `No_giro` (`No_giro`)
) ENGINE=MyISAM DEFAULT CHARSET=utf8 ROW_FORMAT=DYNAMIC;
SET character_set_client = @saved_cs_client;

--
-- Dumping data for table `giro`
--

LOCK TABLES `giro` WRITE;
/*!40000 ALTER TABLE `giro` DISABLE KEYS */;
/*!40000 ALTER TABLE `giro` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `gudang`
--

DROP TABLE IF EXISTS `gudang`;
SET @saved_cs_client     = @@character_set_client;
SET character_set_client = utf8;
CREATE TABLE `gudang` (
  `kode_brg` varchar(50) NOT NULL default '',
  `stok_gudang` double(10,2) default NULL,
  PRIMARY KEY  (`kode_brg`)
) ENGINE=MyISAM DEFAULT CHARSET=utf8;
SET character_set_client = @saved_cs_client;

--
-- Dumping data for table `gudang`
--

LOCK TABLES `gudang` WRITE;
/*!40000 ALTER TABLE `gudang` DISABLE KEYS */;
INSERT INTO `gudang` VALUES ('111',0.00),('3000',0.00),('8886022910266',700.00),('8886022930240',456.00),('8991002101104',230.00),('8991002101746',0.00),('9970001990039',0.00),('Br0001',0.00),('cfdsfewferf',0.00),('dasdsad',0.00),('dfdsfdsf',0.00),('Mk0032',0.00),('Br0005',0.00),('Br0006',0.00),('Br0007',0.00);
/*!40000 ALTER TABLE `gudang` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `harga`
--

DROP TABLE IF EXISTS `harga`;
SET @saved_cs_client     = @@character_set_client;
SET character_set_client = utf8;
CREATE TABLE `harga` (
  `kode_brg` varchar(50) NOT NULL default '',
  `Deskripsi` varchar(255) default NULL,
  `harga` double(19,4) default NULL
) ENGINE=MyISAM DEFAULT CHARSET=utf8;
SET character_set_client = @saved_cs_client;

--
-- Dumping data for table `harga`
--

LOCK TABLES `harga` WRITE;
/*!40000 ALTER TABLE `harga` DISABLE KEYS */;
INSERT INTO `harga` VALUES ('333','baju',0.0000),('666','emas',0.0000),('21312321','hp',0.0000),('222','kursi',0.0000),('111','Meja  xx ',20000.0000);
/*!40000 ALTER TABLE `harga` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `hutang`
--

DROP TABLE IF EXISTS `hutang`;
SET @saved_cs_client     = @@character_set_client;
SET character_set_client = utf8;
CREATE TABLE `hutang` (
  `Tanggal` date default NULL,
  `No_pembelian` varchar(20) default NULL,
  `Jumlah_hutang` double(19,4) default NULL,
  `Jumlah_byr` double(19,4) default NULL,
  `Jatuh_tempo` date default NULL,
  `Id_supplier` varchar(20) default NULL,
  KEY `Id_supplier` (`Id_supplier`),
  KEY `Pembelianhutang` (`No_pembelian`)
) ENGINE=MyISAM DEFAULT CHARSET=utf8 ROW_FORMAT=DYNAMIC;
SET character_set_client = @saved_cs_client;

--
-- Dumping data for table `hutang`
--

LOCK TABLES `hutang` WRITE;
/*!40000 ALTER TABLE `hutang` DISABLE KEYS */;
/*!40000 ALTER TABLE `hutang` ENABLE KEYS */;
UNLOCK TABLES;

/*!50003 SET @SAVE_SQL_MODE=@@SQL_MODE*/;

DELIMITER ;;
/*!50003 SET SESSION SQL_MODE="NO_AUTO_CREATE_USER,NO_ENGINE_SUBSTITUTION" */;;
/*!50003 CREATE */ /*!50017 DEFINER=`root`@`localhost` */ /*!50003 TRIGGER `thutang` BEFORE INSERT ON `hutang` FOR EACH ROW BEGIN
update tblsupplier set jumlah_hutang=jumlah_hutang+new.jumlah_hutang where id_supplier=new.id_supplier;
END */;;

/*!50003 SET SESSION SQL_MODE="NO_AUTO_CREATE_USER,NO_ENGINE_SUBSTITUTION" */;;
/*!50003 CREATE */ /*!50017 DEFINER=`root`@`localhost` */ /*!50003 TRIGGER `thutang2` AFTER DELETE ON `hutang` FOR EACH ROW BEGIN
update tblsupplier set jumlah_hutang=jumlah_hutang-(old.jumlah_hutang-old.jumlah_byr) where id_supplier=old.id_supplier;
END */;;

DELIMITER ;
/*!50003 SET SESSION SQL_MODE=@SAVE_SQL_MODE*/;

--
-- Table structure for table `jurnal`
--

DROP TABLE IF EXISTS `jurnal`;
SET @saved_cs_client     = @@character_set_client;
SET character_set_client = utf8;
CREATE TABLE `jurnal` (
  `tanggal` date default NULL,
  `no_akun` varchar(20) default NULL,
  `keterangan` varchar(255) default NULL,
  `debet` double(19,2) default NULL,
  `kredit` double(19,2) default NULL,
  `no_transaksi` varchar(20) default NULL,
  `nmr` int(15) NOT NULL auto_increment,
  `keterangan2` varchar(255) default NULL,
  PRIMARY KEY  (`nmr`)
) ENGINE=MyISAM AUTO_INCREMENT=121 DEFAULT CHARSET=latin1 ROW_FORMAT=DYNAMIC;
SET character_set_client = @saved_cs_client;

--
-- Dumping data for table `jurnal`
--

LOCK TABLES `jurnal` WRITE;
/*!40000 ALTER TABLE `jurnal` DISABLE KEYS */;
INSERT INTO `jurnal` VALUES ('2013-08-16','4.1','KAS',200000.00,0.00,'By130816001',117,''),('2013-08-16','2.1.','  asds',0.00,200000.00,'By130816001',118,''),('2013-08-17','2.3.','Gaji pegawai',2000000.00,0.00,'Bx130817001',119,''),('2013-08-17','4.1','  Kas',0.00,2000000.00,'Bx130817001',120,'');
/*!40000 ALTER TABLE `jurnal` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `keuangan`
--

DROP TABLE IF EXISTS `keuangan`;
SET @saved_cs_client     = @@character_set_client;
SET character_set_client = utf8;
CREATE TABLE `keuangan` (
  `Tanggal` date default NULL,
  `Keterangan` varchar(255) default NULL,
  `Pemasukan` double(19,4) default NULL,
  `Pengeluaran` double(19,4) default NULL,
  `Jenis` varchar(255) default NULL,
  `no_bayar` varchar(50) default NULL,
  `no_transaksi` varchar(50) default NULL,
  `kasir` varchar(255) default NULL
) ENGINE=MyISAM DEFAULT CHARSET=utf8 ROW_FORMAT=DYNAMIC;
SET character_set_client = @saved_cs_client;

--
-- Dumping data for table `keuangan`
--

LOCK TABLES `keuangan` WRITE;
/*!40000 ALTER TABLE `keuangan` DISABLE KEYS */;
/*!40000 ALTER TABLE `keuangan` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `lapmember`
--

DROP TABLE IF EXISTS `lapmember`;
SET @saved_cs_client     = @@character_set_client;
SET character_set_client = utf8;
CREATE TABLE `lapmember` (
  `id_pelanggan` varchar(50) NOT NULL default '',
  `nama` varchar(255) default NULL,
  `alamat` varchar(255) default NULL,
  `inten` int(3) default NULL,
  `jumlah` double(19,2) default NULL,
  PRIMARY KEY  (`id_pelanggan`)
) ENGINE=MyISAM DEFAULT CHARSET=utf8;
SET character_set_client = @saved_cs_client;

--
-- Dumping data for table `lapmember`
--

LOCK TABLES `lapmember` WRITE;
/*!40000 ALTER TABLE `lapmember` DISABLE KEYS */;
INSERT INTO `lapmember` VALUES ('CUS0010','Agus Eko','',1,56250.00),('CUS0056','8995075700158','',1,30763.00);
/*!40000 ALTER TABLE `lapmember` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `lapmos`
--

DROP TABLE IF EXISTS `lapmos`;
SET @saved_cs_client     = @@character_set_client;
SET character_set_client = utf8;
CREATE TABLE `lapmos` (
  `Kode_brg` varchar(20) NOT NULL default '',
  `deskripsi` varchar(255) default NULL,
  `kategori` varchar(255) default NULL,
  `satuan` varchar(50) default NULL,
  `item` decimal(5,2) default NULL,
  `untung` double(19,4) default NULL,
  `jual` double(19,4) default NULL
) ENGINE=MyISAM DEFAULT CHARSET=utf8 ROW_FORMAT=DYNAMIC;
SET character_set_client = @saved_cs_client;

--
-- Dumping data for table `lapmos`
--

LOCK TABLES `lapmos` WRITE;
/*!40000 ALTER TABLE `lapmos` DISABLE KEYS */;
INSERT INTO `lapmos` VALUES ('Br0001','Baju tidur','Baju','pcs','75.00',28800.0000,178800.0000),('Br0004','Buku','alat tulis','pcs','40.00',40000.0000,240000.0000),('Br0002','Kue','kue','pcs','30.00',18000.0000,108000.0000),('Br0003','Pulpen','alat tulis','pcs','0.00',0.0000,0.0000);
/*!40000 ALTER TABLE `lapmos` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `paste_errors`
--

DROP TABLE IF EXISTS `paste_errors`;
SET @saved_cs_client     = @@character_set_client;
SET character_set_client = utf8;
CREATE TABLE `paste_errors` (
  `No_pemesanan` longtext,
  `Tanggal_pesan` longtext,
  `Id_supplier` longtext
) ENGINE=MyISAM DEFAULT CHARSET=utf8 ROW_FORMAT=DYNAMIC;
SET character_set_client = @saved_cs_client;

--
-- Dumping data for table `paste_errors`
--

LOCK TABLES `paste_errors` WRITE;
/*!40000 ALTER TABLE `paste_errors` DISABLE KEYS */;
INSERT INTO `paste_errors` VALUES ('PO12051101','5/11/2012','SP0001'),('PO12051102','5/11/2012','SP0001'),('PO12051103','5/11/2012','SP0001'),('PO12051104','5/11/2012','SP0001'),('PO12051105','5/11/2012','SP0001');
/*!40000 ALTER TABLE `paste_errors` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `pelanggan`
--

DROP TABLE IF EXISTS `pelanggan`;
SET @saved_cs_client     = @@character_set_client;
SET character_set_client = utf8;
CREATE TABLE `pelanggan` (
  `Id_pelanggan` varchar(50) NOT NULL default '',
  `Nama` varchar(50) NOT NULL,
  `Kontak_person` varchar(50) default NULL,
  `Alamat` varchar(255) default NULL,
  `Kota` varchar(50) default NULL,
  `Telepon` varchar(50) default NULL,
  `Fax` varchar(50) default NULL,
  `Email` varchar(100) default NULL,
  `jumlah_piutang` double(19,4) default NULL,
  PRIMARY KEY  (`Id_pelanggan`)
) ENGINE=MyISAM DEFAULT CHARSET=utf8 ROW_FORMAT=DYNAMIC;
SET character_set_client = @saved_cs_client;

--
-- Dumping data for table `pelanggan`
--

LOCK TABLES `pelanggan` WRITE;
/*!40000 ALTER TABLE `pelanggan` DISABLE KEYS */;
/*!40000 ALTER TABLE `pelanggan` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `pembelian`
--

DROP TABLE IF EXISTS `pembelian`;
SET @saved_cs_client     = @@character_set_client;
SET character_set_client = utf8;
CREATE TABLE `pembelian` (
  `No_pembelian` varchar(20) NOT NULL default '',
  `No_pemesanan` varchar(20) default NULL,
  `Tanggal_pembelian` date default NULL,
  `Id_supplier` varchar(10) default NULL,
  `Total` double(19,2) default '0.00',
  `Total_diskon` double(19,2) default '0.00',
  `Total_stlh_diskon` double(19,2) default '0.00',
  `keterangan1` varchar(255) default NULL,
  `keterangan2` varchar(255) default NULL,
  `cash` double(19,2) NOT NULL default '0.00',
  `hari` int(4) NOT NULL default '0',
  PRIMARY KEY  (`No_pembelian`)
) ENGINE=MyISAM DEFAULT CHARSET=utf8 ROW_FORMAT=DYNAMIC;
SET character_set_client = @saved_cs_client;

--
-- Dumping data for table `pembelian`
--

LOCK TABLES `pembelian` WRITE;
/*!40000 ALTER TABLE `pembelian` DISABLE KEYS */;
/*!40000 ALTER TABLE `pembelian` ENABLE KEYS */;
UNLOCK TABLES;

/*!50003 SET @SAVE_SQL_MODE=@@SQL_MODE*/;

DELIMITER ;;
/*!50003 SET SESSION SQL_MODE="NO_AUTO_CREATE_USER,NO_ENGINE_SUBSTITUTION" */;;
/*!50003 CREATE */ /*!50017 DEFINER=`root`@`localhost` */ /*!50003 TRIGGER `tpembelian` BEFORE INSERT ON `pembelian` FOR EACH ROW BEGIN

declare disk double(6,5);declare cek double(19,2);



END */;;

/*!50003 SET SESSION SQL_MODE="NO_AUTO_CREATE_USER,NO_ENGINE_SUBSTITUTION" */;;
/*!50003 CREATE */ /*!50017 DEFINER=`root`@`localhost` */ /*!50003 TRIGGER `tpembelian2` BEFORE UPDATE ON `pembelian` FOR EACH ROW BEGIN

declare disk double(6,5);declare cek double(19,2);
declare vhari int;
set vhari=old.hari;
set @vhari=vhari;
set new.Total_stlh_diskon=new.total-new.total_diskon;
if new.cash is not null and new.cash > 0  then
if new.cash<=new.total_stlh_diskon then
insert into keuangan(Tanggal,Keterangan,Pengeluaran,no_transaksi) values(new.Tanggal_pembelian,'Transaksi pembelian',new.cash,new.no_pembelian);
else
insert into keuangan(Tanggal,Keterangan,Pengeluaran,no_transaksi) values(new.Tanggal_pembelian,'Transaksi pembelian',new.total,new.no_pembelian);
end if;
end if;
set cek=(select coalesce(jumlah_hutang,0) from hutang where no_pembelian=new.no_pembelian);
if new.cash<new.total_stlh_diskon then 
if cek!=0 then
delete from hutang where no_pembelian=new.no_pembelian;

end if;
insert into hutang values(new.Tanggal_pembelian,new.no_pembelian,new.total_stlh_diskon-new.cash,0,DATE_ADD(new.Tanggal_pembelian,INTERVAL @vhari DAY),new.id_supplier);
end if;


END */;;

/*!50003 SET SESSION SQL_MODE="NO_AUTO_CREATE_USER,NO_ENGINE_SUBSTITUTION" */;;
/*!50003 CREATE */ /*!50017 DEFINER=`root`@`localhost` */ /*!50003 TRIGGER `tpembelian3` AFTER DELETE ON `pembelian` FOR EACH ROW BEGIN
delete from detilbeli where no_pembelian=old.no_pembelian;
delete from hutang  where no_pembelian=old.no_pembelian;
delete from byr_hutang  where no_pembelian=old.no_pembelian;
delete from keuangan  where no_transaksi=old.no_pembelian;
END */;;

DELIMITER ;
/*!50003 SET SESSION SQL_MODE=@SAVE_SQL_MODE*/;

--
-- Table structure for table `pemesanan`
--

DROP TABLE IF EXISTS `pemesanan`;
SET @saved_cs_client     = @@character_set_client;
SET character_set_client = utf8;
CREATE TABLE `pemesanan` (
  `No_pemesanan` varchar(20) NOT NULL default '',
  `Tanggal_pesan` date default NULL,
  `Id_supplier` varchar(10) default NULL,
  PRIMARY KEY  (`No_pemesanan`),
  KEY `Id_supplier` (`Id_supplier`)
) ENGINE=MyISAM DEFAULT CHARSET=utf8 ROW_FORMAT=DYNAMIC;
SET character_set_client = @saved_cs_client;

--
-- Dumping data for table `pemesanan`
--

LOCK TABLES `pemesanan` WRITE;
/*!40000 ALTER TABLE `pemesanan` DISABLE KEYS */;
/*!40000 ALTER TABLE `pemesanan` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `pengguna`
--

DROP TABLE IF EXISTS `pengguna`;
SET @saved_cs_client     = @@character_set_client;
SET character_set_client = utf8;
CREATE TABLE `pengguna` (
  `Username` varchar(255) character set latin1 NOT NULL default '',
  `Nama` varchar(255) character set latin1 default NULL,
  `Bagian` varchar(50) character set latin1 default NULL,
  `Password` varchar(50) character set latin1 default NULL,
  `cek0` enum('1','0') character set latin1 default NULL,
  `cek1` enum('1','0') character set latin1 default NULL,
  `cek2` enum('1','0') character set latin1 default NULL,
  `cek3` enum('1','0') character set latin1 default NULL,
  `cek4` enum('1','0') character set latin1 default NULL,
  `cek5` enum('1','0') character set latin1 default NULL,
  `cek6` enum('1','0') character set latin1 default NULL,
  `cek7` enum('1','0') character set latin1 default NULL,
  `cek8` enum('1','0') character set latin1 default NULL,
  `cek9` enum('1','0') character set latin1 default NULL,
  `cek10` enum('1','0') character set latin1 default NULL,
  `cek11` enum('1','0') character set latin1 default NULL,
  `cek12` enum('1','0') character set latin1 default NULL,
  `cek13` enum('1','0') character set latin1 default NULL,
  `cek14` enum('1','0') character set latin1 default NULL,
  `cek15` enum('1','0') character set latin1 default NULL,
  `cek16` enum('1','0') character set latin1 default NULL,
  PRIMARY KEY  (`Username`)
) ENGINE=MyISAM DEFAULT CHARSET=utf8 ROW_FORMAT=DYNAMIC;
SET character_set_client = @saved_cs_client;

--
-- Dumping data for table `pengguna`
--

LOCK TABLES `pengguna` WRITE;
/*!40000 ALTER TABLE `pengguna` DISABLE KEYS */;
INSERT INTO `pengguna` VALUES ('Admin','admin','admin','21232f297a57a5a743894a0e4a801fc3','1','1','1','1','1','1','1','1','1','1','1','1','1','1','1','1','1'),('Kasir','Deni','kasir','c7911af3adbd12a035b289556d96470a','0','0','0','0','0','0','1','0','0','0','0','0','0','0','0','0','0');
/*!40000 ALTER TABLE `pengguna` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `penjualan`
--

DROP TABLE IF EXISTS `penjualan`;
SET @saved_cs_client     = @@character_set_client;
SET character_set_client = utf8;
CREATE TABLE `penjualan` (
  `No_penjualan` varchar(50) NOT NULL default '',
  `Tanggal` date default NULL,
  `Jumlah` double(19,2) default NULL,
  `Total_diskon` double(19,4) default NULL,
  `Total` double(19,2) default NULL,
  `kasir` varchar(255) default NULL,
  `Id_pelanggan` varchar(50) default NULL,
  `harga_pokok_jual` double(19,4) default NULL,
  `keterangan1` varchar(255) default NULL,
  `keterangan2` varchar(255) default NULL,
  `ppn` double(19,4) default NULL,
  `no_po` varchar(50) default NULL,
  `hari` smallint(6) default NULL,
  `jenis` varchar(30) default NULL,
  `cash` double(19,4) NOT NULL default '0.0000',
  PRIMARY KEY  (`No_penjualan`)
) ENGINE=MyISAM DEFAULT CHARSET=utf8 ROW_FORMAT=DYNAMIC;
SET character_set_client = @saved_cs_client;

--
-- Dumping data for table `penjualan`
--

LOCK TABLES `penjualan` WRITE;
/*!40000 ALTER TABLE `penjualan` DISABLE KEYS */;
/*!40000 ALTER TABLE `penjualan` ENABLE KEYS */;
UNLOCK TABLES;

/*!50003 SET @SAVE_SQL_MODE=@@SQL_MODE*/;

DELIMITER ;;
/*!50003 SET SESSION SQL_MODE="NO_AUTO_CREATE_USER,NO_ENGINE_SUBSTITUTION" */;;
/*!50003 CREATE */ /*!50017 DEFINER=`root`@`localhost` */ /*!50003 TRIGGER `tpenjualan` BEFORE INSERT ON `penjualan` FOR EACH ROW BEGIN

declare disk double(6,5);declare cek double(19,2);


if new.cash=new.total then
set new.keterangan1='C';set new.keterangan2='L';
end if;
END */;;

/*!50003 SET SESSION SQL_MODE="NO_AUTO_CREATE_USER,NO_ENGINE_SUBSTITUTION" */;;
/*!50003 CREATE */ /*!50017 DEFINER=`root`@`localhost` */ /*!50003 TRIGGER `tpenjualan2` BEFORE UPDATE ON `penjualan` FOR EACH ROW BEGIN

declare disk double(6,5);declare cek double(19,2);
declare vhari int;
set vhari=old.hari;
set @vhari=vhari;
set new.Total=new.jumlah-new.total_diskon;


if new.cash is not null and new.cash > 0  then
if new.cash<=new.total then
insert into keuangan(Tanggal,Keterangan,Pemasukan,no_transaksi) values(new.Tanggal,'Transaksi penjualan',new.cash,new.no_penjualan);
else
insert into keuangan(Tanggal,Keterangan,Pemasukan,no_transaksi) values(new.Tanggal,'Transaksi penjualan',new.total,new.no_penjualan);
end if;
end if;


set cek=(select coalesce(jumlah_piutang,0) from piutang where no_penjualan=new.no_penjualan);
if new.cash<new.total then 
if cek!=0 then
delete from piutang where no_penjualan=new.no_penjualan;

end if;
insert into piutang values(new.Tanggal,new.no_penjualan,new.total-new.cash,0,DATE_ADD(new.Tanggal,INTERVAL @vhari DAY),new.id_pelanggan);
end if;


END */;;

/*!50003 SET SESSION SQL_MODE="NO_AUTO_CREATE_USER,NO_ENGINE_SUBSTITUTION" */;;
/*!50003 CREATE */ /*!50017 DEFINER=`root`@`localhost` */ /*!50003 TRIGGER `tpenjualan3` AFTER DELETE ON `penjualan` FOR EACH ROW BEGIN
delete from detiljual where no_penjualan=old.no_penjualan;
delete from piutang  where no_penjualan=old.no_penjualan;
delete from byr_piutang  where no_penjualan=old.no_penjualan;
delete from keuangan  where no_transaksi=old.no_penjualan;
END */;;

DELIMITER ;
/*!50003 SET SESSION SQL_MODE=@SAVE_SQL_MODE*/;

--
-- Table structure for table `penjualans`
--

DROP TABLE IF EXISTS `penjualans`;
SET @saved_cs_client     = @@character_set_client;
SET character_set_client = utf8;
CREATE TABLE `penjualans` (
  `No_penjualan` varchar(50) NOT NULL default '',
  `Tanggal` date default NULL,
  `Jumlah` double(19,2) default NULL,
  `Total_diskon` double(19,4) default NULL,
  `Total` double(19,2) default NULL,
  `kasir` varchar(255) default NULL,
  `Id_pelanggan` varchar(20) default NULL,
  `harga_pokok_jual` double(19,4) default NULL,
  `keterangan1` varchar(255) default NULL,
  `keterangan2` varchar(255) default NULL,
  `ppn` double(19,4) default NULL,
  `no_po` varchar(50) default NULL,
  `hari` smallint(6) default NULL,
  `jenis` varchar(30) default NULL
) ENGINE=MyISAM DEFAULT CHARSET=utf8 ROW_FORMAT=DYNAMIC;
SET character_set_client = @saved_cs_client;

--
-- Dumping data for table `penjualans`
--

LOCK TABLES `penjualans` WRITE;
/*!40000 ALTER TABLE `penjualans` DISABLE KEYS */;
/*!40000 ALTER TABLE `penjualans` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `piutang`
--

DROP TABLE IF EXISTS `piutang`;
SET @saved_cs_client     = @@character_set_client;
SET character_set_client = utf8;
CREATE TABLE `piutang` (
  `Tanggal` date default NULL,
  `No_penjualan` varchar(20) default NULL,
  `Jumlah_piutang` double(19,4) default NULL,
  `Jumlah_byr` double(19,4) default NULL,
  `Jatuh_tempo` date default NULL,
  `Id_pelanggan` varchar(20) default NULL,
  KEY `Id_supplier` (`Id_pelanggan`),
  KEY `PenjualanPiutang` (`No_penjualan`)
) ENGINE=MyISAM DEFAULT CHARSET=utf8 ROW_FORMAT=DYNAMIC;
SET character_set_client = @saved_cs_client;

--
-- Dumping data for table `piutang`
--

LOCK TABLES `piutang` WRITE;
/*!40000 ALTER TABLE `piutang` DISABLE KEYS */;
/*!40000 ALTER TABLE `piutang` ENABLE KEYS */;
UNLOCK TABLES;

/*!50003 SET @SAVE_SQL_MODE=@@SQL_MODE*/;

DELIMITER ;;
/*!50003 SET SESSION SQL_MODE="NO_AUTO_CREATE_USER,NO_ENGINE_SUBSTITUTION" */;;
/*!50003 CREATE */ /*!50017 DEFINER=`root`@`localhost` */ /*!50003 TRIGGER `tpiutang` BEFORE INSERT ON `piutang` FOR EACH ROW BEGIN
update pelanggan set jumlah_piutang=jumlah_piutang+new.jumlah_piutang where id_pelanggan=new.id_pelanggan;
END */;;

/*!50003 SET SESSION SQL_MODE="NO_AUTO_CREATE_USER,NO_ENGINE_SUBSTITUTION" */;;
/*!50003 CREATE */ /*!50017 DEFINER=`root`@`localhost` */ /*!50003 TRIGGER `tpiutang2` AFTER DELETE ON `piutang` FOR EACH ROW BEGIN
update pelanggan set jumlah_piutang=jumlah_piutang-(old.jumlah_piutang-old.jumlah_byr) where id_pelanggan=old.id_pelanggan;
END */;;

DELIMITER ;
/*!50003 SET SESSION SQL_MODE=@SAVE_SQL_MODE*/;

--
-- Temporary table structure for view `rengking`
--

DROP TABLE IF EXISTS `rengking`;
/*!50001 DROP VIEW IF EXISTS `rengking`*/;
/*!50001 CREATE TABLE `rengking` (
  `kode_brg` varchar(20),
  `deskripsi` varchar(255),
  `kategori` varchar(255),
  `satuan` varchar(50),
  `jum` double(19,2),
  `ttl` double(21,4),
  `utg` double(21,4)
) */;

--
-- Table structure for table `retur_beli`
--

DROP TABLE IF EXISTS `retur_beli`;
SET @saved_cs_client     = @@character_set_client;
SET character_set_client = utf8;
CREATE TABLE `retur_beli` (
  `No_retur` varchar(20) NOT NULL default '',
  `Tanggal` date default NULL,
  `total_brg` double(10,2) default NULL,
  `Total` double(19,4) default NULL,
  PRIMARY KEY  (`No_retur`)
) ENGINE=MyISAM DEFAULT CHARSET=utf8 ROW_FORMAT=DYNAMIC;
SET character_set_client = @saved_cs_client;

--
-- Dumping data for table `retur_beli`
--

LOCK TABLES `retur_beli` WRITE;
/*!40000 ALTER TABLE `retur_beli` DISABLE KEYS */;
/*!40000 ALTER TABLE `retur_beli` ENABLE KEYS */;
UNLOCK TABLES;

/*!50003 SET @SAVE_SQL_MODE=@@SQL_MODE*/;

DELIMITER ;;
/*!50003 SET SESSION SQL_MODE="NO_AUTO_CREATE_USER,NO_ENGINE_SUBSTITUTION" */;;
/*!50003 CREATE */ /*!50017 DEFINER=`root`@`localhost` */ /*!50003 TRIGGER `treturbeli` BEFORE UPDATE ON `retur_beli` FOR EACH ROW BEGIN
delete from keuangan where no_transaksi=new.no_retur;
insert into keuangan(tanggal,keterangan,pemasukan,jenis,no_transaksi) values (new.tanggal,'Retur beli',new.total,'Retur beli',new.no_retur);

END */;;

/*!50003 SET SESSION SQL_MODE="NO_AUTO_CREATE_USER,NO_ENGINE_SUBSTITUTION" */;;
/*!50003 CREATE */ /*!50017 DEFINER=`root`@`localhost` */ /*!50003 TRIGGER `treturbeli2` BEFORE DELETE ON `retur_beli` FOR EACH ROW BEGIN
delete from keuangan where no_transaksi=old.no_retur;
delete from detilreturbeli where no_retur=old.no_retur;
END */;;

DELIMITER ;
/*!50003 SET SESSION SQL_MODE=@SAVE_SQL_MODE*/;

--
-- Table structure for table `retur_jual`
--

DROP TABLE IF EXISTS `retur_jual`;
SET @saved_cs_client     = @@character_set_client;
SET character_set_client = utf8;
CREATE TABLE `retur_jual` (
  `No_retur` varchar(20) NOT NULL default '',
  `Tanggal` date default NULL,
  `total_brg` double(10,2) default NULL,
  `Total` double(19,2) default NULL,
  PRIMARY KEY  (`No_retur`)
) ENGINE=MyISAM DEFAULT CHARSET=utf8 ROW_FORMAT=DYNAMIC;
SET character_set_client = @saved_cs_client;

--
-- Dumping data for table `retur_jual`
--

LOCK TABLES `retur_jual` WRITE;
/*!40000 ALTER TABLE `retur_jual` DISABLE KEYS */;
/*!40000 ALTER TABLE `retur_jual` ENABLE KEYS */;
UNLOCK TABLES;

/*!50003 SET @SAVE_SQL_MODE=@@SQL_MODE*/;

DELIMITER ;;
/*!50003 SET SESSION SQL_MODE="NO_AUTO_CREATE_USER,NO_ENGINE_SUBSTITUTION" */;;
/*!50003 CREATE */ /*!50017 DEFINER=`root`@`localhost` */ /*!50003 TRIGGER `treturjual` BEFORE UPDATE ON `retur_jual` FOR EACH ROW BEGIN
delete from keuangan where no_transaksi=new.no_retur;
insert into keuangan(tanggal,keterangan,pengeluaran,jenis,no_transaksi) values (new.tanggal,'Retur jual',new.total,'Retur jual',new.no_retur);

END */;;

/*!50003 SET SESSION SQL_MODE="NO_AUTO_CREATE_USER,NO_ENGINE_SUBSTITUTION" */;;
/*!50003 CREATE */ /*!50017 DEFINER=`root`@`localhost` */ /*!50003 TRIGGER `treturjual2` BEFORE DELETE ON `retur_jual` FOR EACH ROW BEGIN
delete from keuangan where no_transaksi=old.no_retur;
delete from detilreturjual where no_retur=old.no_retur;
END */;;

DELIMITER ;
/*!50003 SET SESSION SQL_MODE=@SAVE_SQL_MODE*/;

--
-- Table structure for table `sales`
--

DROP TABLE IF EXISTS `sales`;
SET @saved_cs_client     = @@character_set_client;
SET character_set_client = utf8;
CREATE TABLE `sales` (
  `sales_id` varchar(50) NOT NULL default '',
  `Nama_sales` varchar(255) NOT NULL,
  `Alamat_sales` varchar(255) default NULL,
  `telp_sales` varchar(50) default NULL,
  PRIMARY KEY  (`sales_id`)
) ENGINE=MyISAM DEFAULT CHARSET=utf8 ROW_FORMAT=DYNAMIC;
SET character_set_client = @saved_cs_client;

--
-- Dumping data for table `sales`
--

LOCK TABLES `sales` WRITE;
/*!40000 ALTER TABLE `sales` DISABLE KEYS */;
INSERT INTO `sales` VALUES ('Sls001','Sales toko','Cilampeniy','022-91590085'),('Sls002','Ade','ase','ase');
/*!40000 ALTER TABLE `sales` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `satuan`
--

DROP TABLE IF EXISTS `satuan`;
SET @saved_cs_client     = @@character_set_client;
SET character_set_client = utf8;
CREATE TABLE `satuan` (
  `Kode_brg` varchar(20) NOT NULL default '',
  `satuan` varchar(50) NOT NULL default '',
  `konversi` float default NULL,
  `keterangan` varchar(255) default NULL,
  `harga` double(19,4) default NULL,
  PRIMARY KEY  (`Kode_brg`,`satuan`),
  KEY `tblbarangsatuan` (`Kode_brg`)
) ENGINE=MyISAM DEFAULT CHARSET=utf8 ROW_FORMAT=DYNAMIC;
SET character_set_client = @saved_cs_client;

--
-- Dumping data for table `satuan`
--

LOCK TABLES `satuan` WRITE;
/*!40000 ALTER TABLE `satuan` DISABLE KEYS */;
/*!40000 ALTER TABLE `satuan` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `sesuai`
--

DROP TABLE IF EXISTS `sesuai`;
SET @saved_cs_client     = @@character_set_client;
SET character_set_client = utf8;
CREATE TABLE `sesuai` (
  `id` int(7) NOT NULL auto_increment,
  `Tanggal` date default NULL,
  `Jenis` varchar(255) default NULL,
  `kode_brg` varchar(20) default NULL,
  `jumlah` double(10,2) default NULL,
  `alasan` varchar(255) default NULL,
  PRIMARY KEY  (`id`)
) ENGINE=MyISAM AUTO_INCREMENT=6 DEFAULT CHARSET=utf8 ROW_FORMAT=DYNAMIC;
SET character_set_client = @saved_cs_client;

--
-- Dumping data for table `sesuai`
--

LOCK TABLES `sesuai` WRITE;
/*!40000 ALTER TABLE `sesuai` DISABLE KEYS */;
/*!40000 ALTER TABLE `sesuai` ENABLE KEYS */;
UNLOCK TABLES;

/*!50003 SET @SAVE_SQL_MODE=@@SQL_MODE*/;

DELIMITER ;;
/*!50003 SET SESSION SQL_MODE="NO_AUTO_CREATE_USER,NO_ENGINE_SUBSTITUTION" */;;
/*!50003 CREATE */ /*!50017 DEFINER=`root`@`localhost` */ /*!50003 TRIGGER `tsesuaiai` AFTER INSERT ON `sesuai` FOR EACH ROW BEGIN
if new.Jenis='Menambah' then
update tblbarang set stok=stok+new.jumlah where kode_brg=new.kode_brg;
else
update tblbarang set stok=stok-new.jumlah where kode_brg=new.kode_brg;
end if;
END */;;

/*!50003 SET SESSION SQL_MODE="NO_AUTO_CREATE_USER,NO_ENGINE_SUBSTITUTION" */;;
/*!50003 CREATE */ /*!50017 DEFINER=`root`@`localhost` */ /*!50003 TRIGGER `tsesuaiad` AFTER DELETE ON `sesuai` FOR EACH ROW BEGIN
if old.Jenis='Menambah' then
update tblbarang set stok=stok-old.jumlah where kode_brg=old.kode_brg;
else
update tblbarang set stok=stok+old.jumlah where kode_brg=old.kode_brg;
end if;
END */;;

DELIMITER ;
/*!50003 SET SESSION SQL_MODE=@SAVE_SQL_MODE*/;

--
-- Table structure for table `sesuai2`
--

DROP TABLE IF EXISTS `sesuai2`;
SET @saved_cs_client     = @@character_set_client;
SET character_set_client = utf8;
CREATE TABLE `sesuai2` (
  `Tanggal` date default NULL,
  `Jenis` varchar(255) default NULL,
  `kode_brg` varchar(20) default NULL,
  `jumlah` double(10,2) default NULL,
  `alasan` varchar(255) default NULL
) ENGINE=MyISAM DEFAULT CHARSET=utf8 ROW_FORMAT=DYNAMIC;
SET character_set_client = @saved_cs_client;

--
-- Dumping data for table `sesuai2`
--

LOCK TABLES `sesuai2` WRITE;
/*!40000 ALTER TABLE `sesuai2` DISABLE KEYS */;
/*!40000 ALTER TABLE `sesuai2` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `stok`
--

DROP TABLE IF EXISTS `stok`;
SET @saved_cs_client     = @@character_set_client;
SET character_set_client = utf8;
CREATE TABLE `stok` (
  `Kode_brg` varchar(50) default NULL,
  `stok` double(10,2) default NULL
) ENGINE=MyISAM DEFAULT CHARSET=utf8 ROW_FORMAT=DYNAMIC;
SET character_set_client = @saved_cs_client;

--
-- Dumping data for table `stok`
--

LOCK TABLES `stok` WRITE;
/*!40000 ALTER TABLE `stok` DISABLE KEYS */;
/*!40000 ALTER TABLE `stok` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `tblbarang`
--

DROP TABLE IF EXISTS `tblbarang`;
SET @saved_cs_client     = @@character_set_client;
SET character_set_client = utf8;
CREATE TABLE `tblbarang` (
  `Kode_brg` varchar(20) NOT NULL default '',
  `Deskripsi` varchar(255) default NULL,
  `kategori` varchar(255) default '0',
  `Satuan` varchar(50) default NULL,
  `Stok` double(10,2) default NULL,
  `Harga_beli` double(19,2) default NULL,
  `Harga_jual` double(19,2) default '0.00',
  `Harga_grosir` double(19,2) default '0.00',
  `diskon` smallint(6) default '0',
  `Stok_minimal` double(10,2) default '0.00',
  PRIMARY KEY  (`Kode_brg`)
) ENGINE=MyISAM DEFAULT CHARSET=utf8 ROW_FORMAT=DYNAMIC;
SET character_set_client = @saved_cs_client;

--
-- Dumping data for table `tblbarang`
--

LOCK TABLES `tblbarang` WRITE;
/*!40000 ALTER TABLE `tblbarang` DISABLE KEYS */;
/*!40000 ALTER TABLE `tblbarang` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `tblsupplier`
--

DROP TABLE IF EXISTS `tblsupplier`;
SET @saved_cs_client     = @@character_set_client;
SET character_set_client = utf8;
CREATE TABLE `tblsupplier` (
  `Id_supplier` varchar(10) NOT NULL default '',
  `Supplier` varchar(255) NOT NULL,
  `Kontak_person` varchar(255) default NULL,
  `Alamat` varchar(255) default NULL,
  `No_telp` varchar(255) NOT NULL,
  `jumlah_hutang` double(19,4) default NULL,
  PRIMARY KEY  (`Id_supplier`),
  KEY `Id_supplier` (`Id_supplier`)
) ENGINE=MyISAM DEFAULT CHARSET=utf8 ROW_FORMAT=DYNAMIC;
SET character_set_client = @saved_cs_client;

--
-- Dumping data for table `tblsupplier`
--

LOCK TABLES `tblsupplier` WRITE;
/*!40000 ALTER TABLE `tblsupplier` DISABLE KEYS */;
/*!40000 ALTER TABLE `tblsupplier` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Final view structure for view `rengking`
--

/*!50001 DROP TABLE `rengking`*/;
/*!50001 DROP VIEW IF EXISTS `rengking`*/;
/*!50001 CREATE ALGORITHM=UNDEFINED */
/*!50013 DEFINER=`root`@`localhost` SQL SECURITY DEFINER */
/*!50001 VIEW `rengking` AS select `t`.`Kode_brg` AS `kode_brg`,`t`.`Deskripsi` AS `deskripsi`,`t`.`kategori` AS `kategori`,`t`.`Satuan` AS `satuan`,coalesce(sum(`d`.`Jumlah_brg`),0) AS `jum`,coalesce(sum(`d`.`Total`),0) AS `ttl`,coalesce(sum((((`d`.`Harga_jual` - `d`.`Harga_beli`) * `d`.`Jumlah_brg`) - `d`.`diskon`)),0) AS `utg` from (`tblbarang` `t` left join `detiljual` `d` on((`d`.`Kode_brg` = `t`.`Kode_brg`))) group by `t`.`Kode_brg` */;
/*!40103 SET TIME_ZONE=@OLD_TIME_ZONE */;

/*!40101 SET SQL_MODE=@OLD_SQL_MODE */;
/*!40014 SET FOREIGN_KEY_CHECKS=@OLD_FOREIGN_KEY_CHECKS */;
/*!40014 SET UNIQUE_CHECKS=@OLD_UNIQUE_CHECKS */;
/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
/*!40111 SET SQL_NOTES=@OLD_SQL_NOTES */;

-- Dump completed on 2013-09-05  9:05:53