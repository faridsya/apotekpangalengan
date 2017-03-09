TYPE=VIEW
query=select `s`.`Tanggal` AS `tanggal`,`b`.`Kode_brg` AS `Kode_brg`,`b`.`Deskripsi` AS `deskripsi`,coalesce(sum((case when (`s`.`Jenis` = _utf8\'Menambah\') then `s`.`jumlah` end)),0) AS `masuk1`,coalesce(sum((case when (`s`.`Jenis` = _utf8\'Mengurang\') then `s`.`jumlah` end)),0) AS `keluar1`,coalesce(sum((case when (`s`.`Jenis` = _utf8\'Menambah\') then `s`.`jumlah` end)),0) AS `summasuk1` from (`apotek1`.`tblbarang` `b` left join `apotek1`.`sesuai` `s` on((`s`.`kode_brg` = `b`.`Kode_brg`))) group by `b`.`Kode_brg`,`s`.`Tanggal`
md5=9ca3321a46bef7e7a0e185aa8dba94da
updatable=0
algorithm=0
definer_user=root
definer_host=localhost
suid=1
with_check_option=0
revision=1
timestamp=2015-05-03 13:33:21
create-version=1
source=select `s`.`Tanggal` AS `tanggal`,`b`.`Kode_brg` AS `Kode_brg`,`b`.`Deskripsi` AS `deskripsi`,coalesce(sum((case when (`s`.`Jenis` = _utf8\'Menambah\') then `s`.`jumlah` end)),0) AS `masuk1`,coalesce(sum((case when (`s`.`Jenis` = _utf8\'Mengurang\') then `s`.`jumlah` end)),0) AS `keluar1`,coalesce(sum((case when (`s`.`Jenis` = _utf8\'Menambah\') then `s`.`jumlah` end)),0) AS `summasuk1` from (`tblbarang` `b` left join `sesuai` `s` on((`s`.`kode_brg` = `b`.`Kode_brg`))) group by `b`.`Kode_brg`,`s`.`Tanggal`
