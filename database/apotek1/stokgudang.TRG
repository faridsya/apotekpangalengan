TYPE=TRIGGERS
triggers='CREATE DEFINER=`root`@`localhost` TRIGGER `tstokgudangau` AFTER UPDATE ON `stokgudang` FOR EACH ROW BEGIN\nif new.kode_gudang=\'utama\' then\nupdate tblbarang set stok=new.jumlah where kode_brg=new.kode_brg ;\nend if;\nEND'
sql_modes=1342177280
definers='root@localhost'
