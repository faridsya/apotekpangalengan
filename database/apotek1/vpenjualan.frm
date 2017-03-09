TYPE=VIEW
query=select `p`.`Tanggal` AS `tanggal`,`b`.`Kode_brg` AS `kode_brg`,`b`.`Deskripsi` AS `deskripsi`,coalesce(sum(`dj`.`Jumlah_brg`),0) AS `terjual` from (`apotek1`.`tblbarang` `b` left join (`apotek1`.`penjualan` `p` join `apotek1`.`detiljual` `dj` on((`dj`.`No_Penjualan` = `p`.`No_penjualan`))) on((`b`.`Kode_brg` = `dj`.`Kode_brg`))) group by `b`.`Kode_brg`,`p`.`Tanggal`
md5=81fa7d60bf9292e882af19ae328046c2
updatable=0
algorithm=0
definer_user=root
definer_host=localhost
suid=1
with_check_option=0
revision=1
timestamp=2015-05-03 13:33:21
create-version=1
source=select `p`.`Tanggal` AS `tanggal`,`b`.`Kode_brg` AS `kode_brg`,`b`.`Deskripsi` AS `deskripsi`,coalesce(sum(`dj`.`Jumlah_brg`),0) AS `terjual` from (`tblbarang` `b` left join (`penjualan` `p` join `detiljual` `dj` on((`dj`.`No_Penjualan` = `p`.`No_penjualan`))) on((`b`.`Kode_brg` = `dj`.`Kode_brg`))) group by `b`.`Kode_brg`,`p`.`Tanggal`
