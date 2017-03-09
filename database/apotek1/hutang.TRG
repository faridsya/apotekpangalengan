TYPE=TRIGGERS
triggers='CREATE DEFINER=`root`@`localhost` TRIGGER `thutang` BEFORE INSERT ON `hutang` FOR EACH ROW BEGIN\nupdate tblsupplier set jumlah_hutang=jumlah_hutang+new.jumlah_hutang where id_supplier=new.id_supplier;\nEND' 'CREATE DEFINER=`root`@`localhost` TRIGGER `thutang2` AFTER DELETE ON `hutang` FOR EACH ROW BEGIN\nupdate tblsupplier set jumlah_hutang=jumlah_hutang-(old.jumlah_hutang-old.jumlah_byr) where id_supplier=old.id_supplier;\nEND'
sql_modes=1342177280 1342177280
definers='root@localhost' 'root@localhost'
