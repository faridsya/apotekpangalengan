TYPE=VIEW
query=select `v1`.`Kode_brg` AS `kode_brg`,`v1`.`deskripsi` AS `deskripsi`,`v1`.`tanggal` AS `tanggal`,sum((`v1`.`masuk1` - `v1`.`keluar1`)) AS `sum(masuk1)` from `apotek1`.`vsesuai` `v1` where (`v1`.`tanggal` > _utf8\'2014-05-01\') group by `v1`.`Kode_brg`
md5=2972487a431fc5780fe612e72f3e002f
updatable=0
algorithm=0
definer_user=root
definer_host=localhost
suid=1
with_check_option=0
revision=1
timestamp=2015-05-03 13:33:21
create-version=1
source=select `v1`.`Kode_brg` AS `kode_brg`,`v1`.`deskripsi` AS `deskripsi`,`v1`.`tanggal` AS `tanggal`,sum((`v1`.`masuk1` - `v1`.`keluar1`)) AS `sum(masuk1)` from `vsesuai` `v1` where (`v1`.`tanggal` > _utf8\'2014-05-01\') group by `v1`.`Kode_brg`
