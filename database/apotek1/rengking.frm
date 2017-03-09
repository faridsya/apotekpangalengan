TYPE=VIEW
query=select `t`.`Kode_brg` AS `kode_brg`,`t`.`Deskripsi` AS `deskripsi`,`t`.`kategori` AS `kategori`,`t`.`Satuan` AS `satuan`,coalesce(sum(`d`.`Jumlah_brg`),0) AS `jum`,coalesce(sum(`d`.`Total`),0) AS `ttl`,coalesce(sum((((`d`.`Harga_jual` - `d`.`Harga_beli`) * `d`.`Jumlah_brg`) - `d`.`diskon`)),0) AS `utg` from (`apotek1`.`tblbarang` `t` left join `apotek1`.`detiljual` `d` on((`d`.`Kode_brg` = `t`.`Kode_brg`))) group by `t`.`Kode_brg`
md5=f770547a8788017b127b39ab6943f388
updatable=0
algorithm=0
definer_user=root
definer_host=localhost
suid=1
with_check_option=0
revision=1
timestamp=2015-05-03 13:33:21
create-version=1
source=select `t`.`Kode_brg` AS `kode_brg`,`t`.`Deskripsi` AS `deskripsi`,`t`.`kategori` AS `kategori`,`t`.`Satuan` AS `satuan`,coalesce(sum(`d`.`Jumlah_brg`),0) AS `jum`,coalesce(sum(`d`.`Total`),0) AS `ttl`,coalesce(sum((((`d`.`Harga_jual` - `d`.`Harga_beli`) * `d`.`Jumlah_brg`) - `d`.`diskon`)),0) AS `utg` from (`tblbarang` `t` left join `detiljual` `d` on((`d`.`Kode_brg` = `t`.`Kode_brg`))) group by `t`.`Kode_brg`
