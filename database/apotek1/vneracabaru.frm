TYPE=VIEW
query=select `a`.`no_akun` AS `no_akun`,`a`.`nama_akun` AS `nama_akun`,`a`.`jns` AS `jns`,`a`.`jns2` AS `jns2`,coalesce(sum((case when (`a`.`jns` = _utf8\'1\') then (`j`.`debet` - `j`.`kredit`) else (`j`.`kredit` - `j`.`debet`) end)),0) AS `saldo`,`j`.`no_transaksi` AS `no_transaksi`,year(`j`.`tanggal`) AS `tahun` from (`apotek1`.`akun` `a` left join `apotek1`.`jurnal` `j` on((`a`.`no_akun` = `j`.`no_akun`))) group by `a`.`no_akun`,year(`j`.`tanggal`)
md5=0866f9685ce42e5d014852f03efaca2b
updatable=0
algorithm=0
definer_user=root
definer_host=localhost
suid=1
with_check_option=0
revision=1
timestamp=2015-05-03 13:33:21
create-version=1
source=select `a`.`no_akun` AS `no_akun`,`a`.`nama_akun` AS `nama_akun`,`a`.`jns` AS `jns`,`a`.`jns2` AS `jns2`,coalesce(sum((case when (`a`.`jns` = _utf8\'1\') then (`j`.`debet` - `j`.`kredit`) else (`j`.`kredit` - `j`.`debet`) end)),0) AS `saldo`,`j`.`no_transaksi` AS `no_transaksi`,year(`j`.`tanggal`) AS `tahun` from (`akun` `a` left join `jurnal` `j` on((`a`.`no_akun` = `j`.`no_akun`))) group by `a`.`no_akun`,year(`j`.`tanggal`)
