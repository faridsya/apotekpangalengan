TYPE=TRIGGERS
triggers='CREATE DEFINER=`root`@`localhost` TRIGGER `tpiutang` BEFORE INSERT ON `piutang` FOR EACH ROW BEGIN\nupdate pelanggan set jumlah_piutang=jumlah_piutang+new.jumlah_piutang where id_pelanggan=new.id_pelanggan;\nEND' 'CREATE DEFINER=`root`@`localhost` TRIGGER `tpiutang2` AFTER DELETE ON `piutang` FOR EACH ROW BEGIN\nupdate pelanggan set jumlah_piutang=jumlah_piutang-old.jumlah_piutang where id_pelanggan=old.id_pelanggan;\nEND'
sql_modes=1342177280 1342177280
definers='root@localhost' 'root@localhost'
