TYPE=VIEW
query=select coalesce(sum((case when ((month(`j`.`tanggal`) = _utf8\'9\') and (year(`j`.`tanggal`) = _utf8\'2014\') and ((`a`.`jns2` = _utf8\'pendapatan\') or (`a`.`jns2` = _utf8\'beban\'))) then (`j`.`kredit` - `j`.`debet`) end)),0) AS `untung1`,coalesce(sum((case when ((month(`j`.`tanggal`) = _utf8\'10\') and (year(`j`.`tanggal`) = _utf8\'2014\') and ((`a`.`jns2` = _utf8\'pendapatan\') or (`a`.`jns2` = _utf8\'beban\'))) then (`j`.`kredit` - `j`.`debet`) end)),0) AS `untung2` from (`apotek1`.`jurnal` `j` join `apotek1`.`akun` `a` on((`a`.`no_akun` = `j`.`no_akun`)))
md5=45395769399482139a7f45e3183c30f8
updatable=0
algorithm=0
definer_user=root
definer_host=localhost
suid=1
with_check_option=0
revision=1
timestamp=2015-05-03 13:33:21
create-version=1
source=select coalesce(sum((case when ((month(`j`.`tanggal`) = _utf8\'9\') and (year(`j`.`tanggal`) = _utf8\'2014\') and ((`a`.`jns2` = _utf8\'pendapatan\') or (`a`.`jns2` = _utf8\'beban\'))) then (`j`.`kredit` - `j`.`debet`) end)),0) AS `untung1`,coalesce(sum((case when ((month(`j`.`tanggal`) = _utf8\'10\') and (year(`j`.`tanggal`) = _utf8\'2014\') and ((`a`.`jns2` = _utf8\'pendapatan\') or (`a`.`jns2` = _utf8\'beban\'))) then (`j`.`kredit` - `j`.`debet`) end)),0) AS `untung2` from (`jurnal` `j` join `akun` `a` on((`a`.`no_akun` = `j`.`no_akun`)))
