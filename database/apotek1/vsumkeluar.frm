TYPE=VIEW
query=select `v2`.`kode_brg` AS `kode_brg`,`v2`.`deskripsi` AS `deskripsi`,`v2`.`tanggal` AS `tanggal`,sum(`v2`.`terjual`) AS `sum(terjual)` from `apotek1`.`vpenjualan` `v2` where (`v2`.`tanggal` > _utf8\'2014-05-01\') group by `v2`.`kode_brg`
md5=1c67b349e5c2ef68f91aaba400336da4
updatable=0
algorithm=0
definer_user=root
definer_host=localhost
suid=1
with_check_option=0
revision=1
timestamp=2015-05-03 13:33:21
create-version=1
source=select `v2`.`kode_brg` AS `kode_brg`,`v2`.`deskripsi` AS `deskripsi`,`v2`.`tanggal` AS `tanggal`,sum(`v2`.`terjual`) AS `sum(terjual)` from `vpenjualan` `v2` where (`v2`.`tanggal` > _utf8\'2014-05-01\') group by `v2`.`kode_brg`
