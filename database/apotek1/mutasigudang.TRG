TYPE=TRIGGERS
triggers='CREATE DEFINER=`root`@`localhost` TRIGGER `tmutasigudangad` AFTER DELETE ON `mutasigudang` FOR EACH ROW BEGIN\nupdate stokgudang set jumlah=jumlah+old.jumlah where kode_brg=old.kode_brg and kode_gudang=old.dari;\nupdate stokgudang set jumlah=jumlah-old.jumlah where kode_brg=old.kode_brg and kode_gudang=old.ke;\nEND' 'CREATE DEFINER=`root`@`localhost` TRIGGER `tmutasigudangai` AFTER INSERT ON `mutasigudang` FOR EACH ROW BEGIN\nupdate stokgudang set jumlah=jumlah-new.jumlah where kode_brg=new.kode_brg and kode_gudang=new.dari;\nupdate stokgudang set jumlah=jumlah+new.jumlah where kode_brg=new.kode_brg and kode_gudang=new.ke;\nEND'
sql_modes=1342177280 1342177280
definers='root@localhost' 'root@localhost'
