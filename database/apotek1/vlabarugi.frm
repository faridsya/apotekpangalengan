TYPE=VIEW
query=select `vlaba`.`no_akun` AS `no_akun`,`vlaba`.`jns` AS `jns`,`vlaba`.`jns2` AS `jns2`,`vlaba`.`urut` AS `urut`,`vlaba`.`nama_akun` AS `nama_akun`,coalesce(sum((case when (year(`vlaba`.`tanggal`) = _utf8\'2014\') then `vlaba`.`saldo` end)),0) AS `saldo1`,coalesce(sum((case when (year(`vlaba`.`tanggal`) = _utf8\'2015\') then `vlaba`.`saldo` end)),0) AS `saldo2` from `apotek1`.`vlaba` group by `vlaba`.`no_akun`
md5=0bb24a786734a53bdf9fae13c8961d7a
updatable=0
algorithm=0
definer_user=root
definer_host=localhost
suid=1
with_check_option=0
revision=1
timestamp=2015-05-05 13:59:14
create-version=1
source=select `vlaba`.`no_akun` AS `no_akun`,`vlaba`.`jns` AS `jns`,`vlaba`.`jns2` AS `jns2`,`vlaba`.`urut` AS `urut`,`vlaba`.`nama_akun` AS `nama_akun`,coalesce(sum((case when (year(`vlaba`.`tanggal`) = _utf8\'2014\') then `vlaba`.`saldo` end)),0) AS `saldo1`,coalesce(sum((case when (year(`vlaba`.`tanggal`) = _utf8\'2015\') then `vlaba`.`saldo` end)),0) AS `saldo2` from apotek1.`vlaba` group by apotek1.`vlaba`.`no_akun`
