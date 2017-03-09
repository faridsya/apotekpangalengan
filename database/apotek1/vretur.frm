TYPE=VIEW
query=select `r`.`Tanggal` AS `tanggal`,`b`.`Kode_brg` AS `kode_brg`,`b`.`Deskripsi` AS `deskripsi`,coalesce(sum(`dr`.`Jumlah`),0) AS `retur` from (`apotek1`.`tblbarang` `b` left join (`apotek1`.`retur_jual` `r` join `apotek1`.`detilreturjual` `dr` on((`dr`.`No_retur` = `r`.`No_retur`))) on((`b`.`Kode_brg` = `dr`.`Kode_brg`))) group by `b`.`Kode_brg`,`r`.`Tanggal`
md5=2884597af86d173fd66ae72330737aad
updatable=0
algorithm=0
definer_user=root
definer_host=localhost
suid=1
with_check_option=0
revision=1
timestamp=2015-05-03 13:33:21
create-version=1
source=select `r`.`Tanggal` AS `tanggal`,`b`.`Kode_brg` AS `kode_brg`,`b`.`Deskripsi` AS `deskripsi`,coalesce(sum(`dr`.`Jumlah`),0) AS `retur` from (`tblbarang` `b` left join (`retur_jual` `r` join `detilreturjual` `dr` on((`dr`.`No_retur` = `r`.`No_retur`))) on((`b`.`Kode_brg` = `dr`.`Kode_brg`))) group by `b`.`Kode_brg`,`r`.`Tanggal`
