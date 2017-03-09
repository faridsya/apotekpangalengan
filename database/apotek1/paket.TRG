TYPE=TRIGGERS
triggers='CREATE DEFINER=`root`@`localhost` TRIGGER `tpaketad` AFTER DELETE ON `paket` FOR EACH ROW BEGIN\ndelete from paket_detil where kode_pkt=old.kode_pkt;\nEND'
sql_modes=1342177280
definers='root@localhost'
