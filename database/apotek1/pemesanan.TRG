TYPE=TRIGGERS
triggers='CREATE DEFINER=`root`@`localhost` TRIGGER `tpemesanan` AFTER DELETE ON `pemesanan` FOR EACH ROW BEGIN\ndelete from detilpesan where no_pemesanan=old.no_pemesanan;\nEND'
sql_modes=1342177280
definers='root@localhost'
