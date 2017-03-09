TYPE=VIEW
query=select `apotek1`.`vneracabaruakhir`.`no_akun` AS `no_akun`,`apotek1`.`vneracabaruakhir`.`jns` AS `jns`,`apotek1`.`vneracabaruakhir`.`jns2` AS `jns2`,`apotek1`.`vneracabaruakhir`.`urut` AS `urut`,`apotek1`.`vneracabaruakhir`.`nama_akun` AS `nama_akun`,coalesce(sum((case when ((month(`apotek1`.`vneracabaruakhir`.`tanggal`) = _utf8\'9\') and (year(`apotek1`.`vneracabaruakhir`.`tanggal`) = _utf8\'2014\')) then `apotek1`.`vneracabaruakhir`.`saldo` end)),0) AS `saldo1`,coalesce(sum((case when ((month(`apotek1`.`vneracabaruakhir`.`tanggal`) = _utf8\'10\') and (year(`apotek1`.`vneracabaruakhir`.`tanggal`) = _utf8\'2014\')) then `apotek1`.`vneracabaruakhir`.`saldo` end)),0) AS `saldo2` from `apotek1`.`vneracabaruakhir` group by `apotek1`.`vneracabaruakhir`.`no_akun`
md5=dca039ad3536837990da9c6c33928b46
updatable=0
algorithm=0
definer_user=root
definer_host=localhost
suid=1
with_check_option=0
revision=1
timestamp=2015-05-03 13:33:21
create-version=1
source=select `vneracabaruakhir`.`no_akun` AS `no_akun`,`vneracabaruakhir`.`jns` AS `jns`,`vneracabaruakhir`.`jns2` AS `jns2`,`vneracabaruakhir`.`urut` AS `urut`,`vneracabaruakhir`.`nama_akun` AS `nama_akun`,coalesce(sum((case when ((month(`vneracabaruakhir`.`tanggal`) = _utf8\'9\') and (year(`vneracabaruakhir`.`tanggal`) = _utf8\'2014\')) then `vneracabaruakhir`.`saldo` end)),0) AS `saldo1`,coalesce(sum((case when ((month(`vneracabaruakhir`.`tanggal`) = _utf8\'10\') and (year(`vneracabaruakhir`.`tanggal`) = _utf8\'2014\')) then `vneracabaruakhir`.`saldo` end)),0) AS `saldo2` from `vneracabaruakhir` group by `vneracabaruakhir`.`no_akun`
