TYPE=TRIGGERS
triggers='CREATE DEFINER=`root`@`localhost` TRIGGER `ttservisdtl2ad` AFTER DELETE ON `tservis_dtl2` FOR EACH ROW BEGIN\nif old.stts_brg=\'out\' then\nupdate tblbarang set stok=stok+old.jum_brg where kode_brg=old.kode_brg;\nend if;\nEND' 'CREATE DEFINER=`root`@`localhost` TRIGGER `ttservisdtl2au` AFTER UPDATE ON `tservis_dtl2` FOR EACH ROW BEGIN\nif old.stts_brg!=\'out\' and new.stts_brg=\'out\' then\nupdate tblbarang set stok=stok-new.jum_brg where kode_brg=new.kode_brg;\nend if;\nEND'
sql_modes=1342177280 1342177280
definers='root@localhost' 'root@localhost'
