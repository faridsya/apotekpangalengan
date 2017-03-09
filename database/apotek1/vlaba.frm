TYPE=VIEW
query=select `a`.`no_akun` AS `no_akun`,`a`.`nama_akun` AS `nama_akun`,`a`.`jns` AS `jns`,`a`.`jns2` AS `jns2`,`a`.`urut` AS `urut`,coalesce(sum((case when (`a`.`jns` = _utf8\'1\') then (`j`.`debet` - `j`.`kredit`) else (`j`.`kredit` - `j`.`debet`) end)),0) AS `saldo`,`j`.`tanggal` AS `tanggal` from (`apotek1`.`akun` `a` left join `apotek1`.`jurnal` `j` on((`a`.`no_akun` = `j`.`no_akun`))) where ((`a`.`jns2` = _utf8\'beban\') or (`a`.`jns2` = _utf8\'pendapatan\')) group by `a`.`no_akun`,`j`.`tanggal`
md5=1e4d78274a48ce57cb17a92d4b9eea47
updatable=0
algorithm=0
definer_user=root
definer_host=localhost
suid=1
with_check_option=0
revision=1
timestamp=2015-05-03 13:33:21
create-version=1
source=select `a`.`no_akun` AS `no_akun`,`a`.`nama_akun` AS `nama_akun`,`a`.`jns` AS `jns`,`a`.`jns2` AS `jns2`,`a`.`urut` AS `urut`,coalesce(sum((case when (`a`.`jns` = _utf8\'1\') then (`j`.`debet` - `j`.`kredit`) else (`j`.`kredit` - `j`.`debet`) end)),0) AS `saldo`,`j`.`tanggal` AS `tanggal` from (`akun` `a` left join `jurnal` `j` on((`a`.`no_akun` = `j`.`no_akun`))) where ((`a`.`jns2` = _utf8\'beban\') or (`a`.`jns2` = _utf8\'pendapatan\')) group by `a`.`no_akun`,`j`.`tanggal`
