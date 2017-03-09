TYPE=VIEW
query=select `a`.`no_akun` AS `no_akun`,`a`.`nama_akun` AS `nama_akun`,`a`.`jns` AS `jns`,`a`.`jns2` AS `jns2`,coalesce(sum((case when (`a`.`jns` = _utf8\'1\') then (`j`.`debet` - `j`.`kredit`) else (`j`.`kredit` - `j`.`debet`) end)),0) AS `saldo`,`j`.`no_transaksi` AS `no_transaksi` from (`apotek1`.`akun` `a` left join `apotek1`.`jurnal` `j` on(((`a`.`no_akun` = `j`.`no_akun`) and (`j`.`keterangan2` = _utf8\'neraca awal\') and (year(`j`.`tanggal`) = _utf8\'2014\')))) group by `a`.`no_akun`
md5=be3cf354859e41f76cf36afee6f0552a
updatable=0
algorithm=0
definer_user=root
definer_host=localhost
suid=1
with_check_option=0
revision=1
timestamp=2015-05-03 13:33:21
create-version=1
source=select `a`.`no_akun` AS `no_akun`,`a`.`nama_akun` AS `nama_akun`,`a`.`jns` AS `jns`,`a`.`jns2` AS `jns2`,coalesce(sum((case when (`a`.`jns` = _utf8\'1\') then (`j`.`debet` - `j`.`kredit`) else (`j`.`kredit` - `j`.`debet`) end)),0) AS `saldo`,`j`.`no_transaksi` AS `no_transaksi` from (`akun` `a` left join `jurnal` `j` on(((`a`.`no_akun` = `j`.`no_akun`) and (`j`.`keterangan2` = _utf8\'neraca awal\') and (year(`j`.`tanggal`) = _utf8\'2014\')))) group by `a`.`no_akun`
