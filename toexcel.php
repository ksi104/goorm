<?
	include "./common/lib.php";
	$db=dbcon();

	if($_SESSION['mseq']!='') {
		$mseq=$_SESSION['mseq'];
	}else {
		$mseq="0";
	}

	if($mseq=="0") {
		msg("로그인하시기 바랍니다.","js_parent_url","./","");
		exit;
	}

	$menu_seq=$_GET['menu_seq'];
	$menu_info=menu_info($menu_seq,'');
	$menu_tree=explode(",",$menu_info['menu_tree']);
	$permission=mem_permission($mseq,$menu_seq,'edit');
	if($permission=="N") {		
		$alert_txt="권한이 없습니다.";
		
		msg($alert_txt,'','','');
		exit;
	}

require_once './Classes/PHPExcel.php';
function column_char($i) { return chr( 65 + $i ); }

$addquery="";

if($_GET['mode']=="member_list") {
	$filename=date("YmdHis")."회원리스트.xlsx";

	$headers = array('소속','부서','직책','이름','이메일','전화번호','휴대폰','가입일','최근접속일','회원등급');

	if(!$_GET['tab'] || $_GET['tab']=='') $tab='T';
	else $tab=$_GET['tab'];

	if($_GET['searchTxt']) {
		if($_GET['searchType']=='name') {
			$addquery.=" and a.member_nm='".$_GET['searchTxt']."'";
		}else if($_GET['searchType']=='email') {
			$addquery.=" and a.member_id='".$_GET['searchTxt']."'";
		}else if($_GET['searchType']=='sosok') {
			$addquery.=" and c.company_nick_ko like '%".$_GET['searchTxt']."%'";
		}else if($_GET['searchType']=='permission') {
			$addquery.=" and e.permission_nm like '%".$_GET['searchTxt']."%'";
		}else {
			$addquery.=" and (a.member_nm='".$_GET['searchTxt']."' or a.member_id='".$_GET['searchTxt']."' or c.company_nick_ko like '%".$_GET['searchTxt']."%' or e.permission_nm like '%".$_GET['searchTxt']."%')";
		}
		$addstring.="&searchTxt=".urlencode($_GET['searchTxt'])."&searchType=".urlencode($_GET['searchType']);
	}

	if(!$_GET['sortcol']) {
		$sortcol='date';
	}else {
		$sortcol=$_GET['sortcol'];
		$addstring.="&sortcol=".$sortcol;
	}

	if(!$_GET['sortmethod']) {
		$sortmethod='desc';
	}else {
		$sortmethod=$_GET['sortmethod'];
		$addstring.="&sortmethod=".$sortmethod;
	}

	if($tab!='T') $addquery.=" and a.member_state='".$tab."'";

	if($sortcol=="sosok") $orderby="c.company_nick_ko";
	else if($sortcol=="name") $orderby="a.member_nm";
	else if($sortcol=="email") $orderby="a.member_id";
	else if($sortcol=="hp") $orderby="a.member_hp";
	else if($sortcol=="permission") $orderby="permission_nm";
	else $orderby="a.member_seq";

	$query="select a.*, b.fk_company, c.company_nick_ko, if(a.member_state='N','탈퇴회원',if(a.member_state='R','인증대기',if(a.member_state='D','휴면계정',e.permission_nm))) as permission_nm from TB_MEMBER a left join TB_AFFILIATION b on a.member_seq=b.fk_member and b.affiliation_flag='Y' left join TB_COMPANY c on b.fk_company=c.company_seq left join TB_MEMBER_PERMISSION d on a.member_seq=d.fk_member left join TB_PERMISSION_NM e on d.fk_permission_nm=e.permission_nm_seq where a.member_reg_method='REG'".$addquery." order by ".$orderby." ".$sortmethod;
	
	$result=mysqli_query($db,$query);
	$i=0;
	while($row=mysqli_fetch_array($result)) {
		//array('소속','부서','직책','이름','이메일','전화번호','휴대폰','가입일','최근접속일','회원등급');

		$tel_array=explode("-",$row['affiliation_tel']);
		$tel="";
		if($tel_array[0]!='' && $tel_array[1]!='' && $tel_array[2]!='' && $tel_array[3]!=''){
			$tel=$row['affiliation_tel'];
			if(substr($tel,0,3)=="82-") {		
				$tel="0".substr($tel,3);
			}
		}

		$hp_array=explode("-",$row['member_hp']);
		$hp="";
		if($hp_array[0]!='' && $hp_array[1]!='' && $hp_array[2]!='' && $hp_array[3]!=''){
			$hp=$row['member_hp'];
			if(substr($hp,0,3)=="82-") {		
				$hp="0".substr($hp,3);
			}
		}

		$j=0;
		$rows[$i][$j]=stripslashes_deep($row['company_nick_ko']);
		$j++;
		$rows[$i][$j]=stripslashes_deep($row['affiliation_part']);
		$j++;
		$rows[$i][$j]=stripslashes_deep($row['affiliation_position']);
		$j++;
		$rows[$i][$j]=stripslashes_deep($row['member_nm']);
		$j++;
		$rows[$i][$j]=stripslashes_deep($row['member_id']);
		$j++;
		$rows[$i][$j]=$tel;
		$j++;
		$rows[$i][$j]=$hp;
		$j++;
		$rows[$i][$j]=substr($row['member_reg_dt'],0,10);
		$j++;
		$rows[$i][$j]=($row['member_last_login']=='0000-00-00 00:00:00')?'':substr($row['member_last_login'],0,10);
		$j++;
		$rows[$i][$j]=$row['permission_nm'];
		$j++;

		$i++;
	}
	//print_r($rows);
	$num=$i+1;
	$data = array_merge(array($headers), $rows);

	$widths = array(35,25,25,15,35,15,15,20,20,15);
	$header_bgcolor = 'eeeeee';


}else if($_GET['mode']=="nomember_list") {
	$filename=date("YmdHis")."비회원메일리스트.xlsx";

	$headers = array('이름','조직명','이메일','등록일','수신동의');

	$searchType_txt="이름";
	if($_GET['searchTxt']) {
		if($_GET['searchType']=="email") {
			$searchType_txt="이메일";
			$addquery.=" and USEREMAIL='".$_GET['searchTxt']."'";
		}else if($_GET['searchType']=="org") {
			$searchType_txt="조직";
			$addquery.=" and ORGNAME like '%".$_GET['searchTxt']."%'";
		}else {
			$searchType_txt="이름";
			$addquery.=" and USERENAME='".$_GET['searchTxt']."'";
		}

		$addstring.="&searchTxt=".urlencode($_GET['searchTxt'])."&searchType=".urlencode($_GET['searchType']);
	}

	if(!$_GET['sortcol']) {
		$sortcol='date';
	}else {
		$sortcol=$_GET['sortcol'];
		$addstring.="&sortcol=".$sortcol;
	}

	if(!$_GET['sortmethod']) {
		$sortmethod='desc';
	}else {
		$sortmethod=$_GET['sortmethod'];
		$addstring.="&sortmethod=".$sortmethod;
	}

	if($sortcol=="name") $orderby="USERNAME";
	else if($sortcol=="email") $orderby="USERMAIL";
	else $orderby="seq";
	
	if($addquery!='') {
		$addquery=" where ".substr($addquery,4);
	}

	$query="select*from TB_MAILLIST".$addquery." order by ".$orderby." ".$sortmethod;
	$result=mysqli_query($db,$query);
	$i=0;
	while($row=mysqli_fetch_array($result)) {

		$j=0;
		$rows[$i][$j]=$row['USERNAME'];
		$j++;
		$rows[$i][$j]=$row['ORGNAME'];
		$j++;
		$rows[$i][$j]=$row['USERMAIL'];
		$j++;
		$rows[$i][$j]=substr($row['REG_DATE'],0,10);
		$j++;
		$rows[$i][$j]=$row['ISOK'];
		$j++;
		

		$i++;
	}
	//print_r($rows);
	$num=$i+1;
	$data = array_merge(array($headers), $rows);

	$widths = array(15,35,45,15,15);
	$header_bgcolor = 'eeeeee';

}else if($_GET['mode']=="company_list") {
	$filename=date("YmdHis")."기업리스트.xlsx";

	$finance_query1="(select * from TB_FINANCE where (finance_year, fk_company) IN (select max(finance_year) AS finance_year , fk_company from TB_FINANCE where finance_sales > 0 GROUP BY fk_company )) e";
	if($_GET['search']!='') {
		$filter_btn="";
		if($_GET['search']['company_nm']!='') {
			if($addquery=='') $addquery.=" where";
			else $addquery.=" and";
			$addquery.=" (c.company_nm_ko like '%".$_GET['search']['company_nm']."%' or c.company_nick_ko like '%".$_GET['search']['company_nm']."%' or c.company_nm_en like '%".$_GET['search']['company_nm']."%' or c.company_nick_en like '%".$_GET['search']['company_nm']."%')";
			$addstring.="&search[company_nm]=".urlencode($_GET['search']['company_nm']);
			$filter_btn.="<button type='button' class='btn btn-secondary btn-sm' data-field='text|company_nm'>".$_GET['search']['company_nm']." <span>X</span></button>";
		}

		if($_GET['search']['membership']!='') {
			$saddq1="";
			$i=0;
			foreach($_GET['search']['membership'] as $key => $value) {
				
				if($value!='') {
					switch($value) {
						case "00" : $value_txt="비회원사"; break;
						case "01" : $value_txt="회원사"; break;
						default : $value_txt=""; break;
					}
					if($saddq1!='') $saddq1.=" or";
					$saddq1.=" c.company_membership='".$value."'";
					if($value==00) $saddq1.=" or c.company_membership is null";
					$addstring.="&search[membership][".$key."]=".$value;
					$filter_btn.="<button type='button' class='btn btn-secondary btn-sm' data-field='checkbox|membership|".$key."'>".$value_txt." <span>X</span></button>";
					$i++;
				}
			}
			if($saddq1!='') {
				if($addquery=='') $addquery.=" where";
				else $addquery.=" and";
				if($i==1) {
					$addquery.=" (".$saddq1.")";
				}else {
					$addquery.=" (".$saddq1.")";
				}					
			}
		}

		if($_GET['search']['mng_flag']!='') {
			$saddq1="";
			$i=0;
			foreach($_GET['search']['mng_flag'] as $key => $value) {
				
				if($value!='') {
					switch($value) {
						case "Y" : $value_txt="관리기업"; break;
						case "N" : $value_txt="비관리기업"; break;
						default : $value_txt=""; break;
					}
					if($saddq1!='') $saddq1.=" or";
					$saddq1.=" c.mng_flag='".$value."'";
					$addstring.="&search[mng_flag][".$key."]=".$value;
					$filter_btn.="<button type='button' class='btn btn-secondary btn-sm' data-field='checkbox|mng_flag|".$key."'>".$value_txt." <span>X</span></button>";
					$i++;
				}
			}
			if($saddq1!='') {
				if($addquery=='') $addquery.=" where";
				else $addquery.=" and";
				if($i==1) {
					$addquery.=$saddq1;
				}else {
					$addquery.=" (".$saddq1.")";
				}					
			}
		}

		if($_GET['search']['cd_01']!='') {
			$saddq1="";
			$i=0;
			foreach($_GET['search']['cd_01'] as $key => $value) {
				
				if($value!='') {
					$value_txt="";
					$squery1="select*from TB_COMMON_CD where common_cd_seq='".$value."'";
					$sresult1=mysqli_query($db,$squery1);
					$srow1=mysqli_fetch_array($sresult1);
					mysqli_free_result($sresult1);
					$value_txt=$srow1['common_cd'];
					if($saddq1!='') $saddq1.=" or";
					$saddq1.=" c.fk_common_cd_01='".$value."'";
					if($value==5) $saddq1.=" or c.fk_common_cd_01=''";
					$addstring.="&search[cd_01][".$key."]=".$value;
					$filter_btn.="<button type='button' class='btn btn-secondary btn-sm' data-field='checkbox|cd_01|".$key."'>".$value_txt." <span>X</span></button>";
					$i++;
				}
			}
			if($saddq1!='') {
				if($addquery=='') $addquery.=" where";
				else $addquery.=" and";
				if($i==1) {
					$addquery.=" (".$saddq1.")";
				}else {
					$addquery.=" (".$saddq1.")";
				}					
			}
		}

		if($_GET['search']['cd_02']!='') {
			$saddq1="";
			$i=0;
			foreach($_GET['search']['cd_02'] as $key => $value) {
				
				if($value!='') {
					$value_txt="";
					$squery1="select*from TB_COMMON_CD where common_cd_seq='".$value."'";
					$sresult1=mysqli_query($db,$squery1);
					$srow1=mysqli_fetch_array($sresult1);
					mysqli_free_result($sresult1);
					$value_txt=$srow1['common_cd'];
					if($saddq1!='') $saddq1.=" or";
					$saddq1.=" c.fk_common_cd_02 like'%[".$value."]%'";
					$addstring.="&search[cd_02][".$key."]=".$value;
					$filter_btn.="<button type='button' class='btn btn-secondary btn-sm' data-field='checkbox|cd_02|".$key."'>".$value_txt." <span>X</span></button>";
					$i++;
				}
			}
			if($saddq1!='') {
				if($addquery=='') $addquery.=" where";
				else $addquery.=" and";
				if($i==1) {
					$addquery.=$saddq1;
				}else {
					$addquery.=" (".$saddq1.")";
				}					
			}
		}

		if($_GET['search']['setup_dt']!='') {
			$saddq1="";
			$nowY=date("Y");
			if($_GET['search']['setup_dt']=="up") {				
				$sy=($nowY-7)."-12-31";
				$saddq1=" UNIX_TIMESTAMP(c.company_setup_dt) < UNIX_TIMESTAMP('$sy')";
			}else {
				$sy=($nowY-$_GET['search']['setup_dt'])."-01-01";
				$saddq1=" UNIX_TIMESTAMP(c.company_setup_dt) >= UNIX_TIMESTAMP('$sy')";
			}
			if($saddq1!='') {
				if($addquery=='') $addquery.=" where";
				else $addquery.=" and";
				$addquery.=$saddq1;
				$addstring.="&search[setup_dt]=".$_GET['search']['setup_dt'];
				switch($_GET['search']['setup_dt']) {
					case "1" : $value_txt="1년 이내"; $key=0; break;
					case "3" : $value_txt="3년 이내"; $key=1; break;
					case "5" : $value_txt="5년 이내"; $key=2; break;
					case "7" : $value_txt="7년 이내"; $key=3; break;
					default : $value_txt="7년 이상"; $key=4; break;
				}
				$filter_btn.="<button type='button' class='btn btn-secondary btn-sm' data-field='radio|setup_dt|".$key."'>".$value_txt." <span>X</span></button>";		
			}
		}

        $setup_dt_range="";
        $sdr1="";
        $sdr2="";
        if($_GET['search']['setup_dt_range']!='' && $_GET['search']['setup_dt_range']!='Array') {
            $setup_dt_range=explode("~",$_GET['search']['setup_dt_range']);
            $sdr1=$setup_dt_range[0];
            $sdr2=$setup_dt_range[1];
            $saddq1="";
            $sy=$_GET['search']['setup_dt1']."-01-01";
            $ey=$_GET['search']['setup_dt2']."-12-31";
            if($addquery=='') $addquery.=" where";
            else $addquery.=" and";
            $addquery.=" UNIX_TIMESTAMP(c.company_setup_dt) >= UNIX_TIMESTAMP('".$sy."') and UNIX_TIMESTAMP(c.company_setup_dt) <= UNIX_TIMESTAMP('".$ey."')";
            $addstring.="&search[setup_dt_range]=".urlencode($_GET['search']['setup_dt_range']);
            $filter_btn.="<button type='button' class='btn btn-secondary btn-sm' data-field='text|setup_dt_range'>".$_GET['search']['setup_dt_range']." <span>X</span></button>";

        }

        
        if($_GET['search']['sales_year']!='') {
            $finance_query1="(select*from TB_FINANCE where finance_year='".$_GET['search']['sales_year']."') e";
            $addstring.="&search[sales_year]=".$_GET['search']['sales_year'];
        }
        if($_GET['search']['sales']!='') {
            $sales=$_GET['search']['sales']."00";
            if($addquery=='') $addquery.=" where";
            else $addquery.=" and";
            $addquery.=" e.finance_sales >= ".$sales;
            $addstring.="&search[sales]=".$_GET['search']['sales'];
            switch($_GET['search']['sales']) {
                case "10" : $value_txt="10억 이상"; $key=0; break;
                case "20" : $value_txt="20억 이상"; $key=1; break;
                case "30" : $value_txt="30억 이상"; $key=2; break;
                default : $value_txt=""; $key=''; break;
            }
            $filter_btn.="<button type='button' class='btn btn-secondary btn-sm' data-field='radio|sales|".$key."'>".$value_txt." <span>X</span></button>";
        }

        $sales_range="";
        $sales1="";
        $sales2="";
        if($_GET['search']['sales_range']!='' && $_GET['search']['sales_range']!='Array') {
            $sales_range=explode("~",$_GET['search']['sales_range']);
            $sales1=$sales_range[0];
            $sales2=$sales_range[1];
            $saddq1="";
            $s1=$sales1*100;
            $s2=$sales2*100;
            if($addquery=='') $addquery.=" where";
            else $addquery.=" and";
            $addquery.=" e.finance_sales >= ".$s1." and e.finance_sales <=".$s2;
            $addstring.="&search[sales_range]=".urlencode($_GET['search']['sales_range']);
            $filter_btn.="<button type='button' class='btn btn-secondary btn-sm' data-field='text|sales_range'>".$_GET['search']['sales_range']." <span>X</span></button>";

        }

		

		if($_GET['search']['empl']!='') {
			$saddq1="";
			$i=0;
			foreach($_GET['search']['empl'] as $key => $value) {
				
				if($value!='') {
					switch($key) {
						case "v1" : $value_txt="1~10명"; $a=1; $b=10; break;
						case "v2" : $value_txt="11~20명"; $a=11; $b=20; break;
						case "v3" : $value_txt="21~50명"; $a=21; $b=50; break;
						case "v4" : $value_txt="51~100명"; $a=51; $b=100; break;
						case "v5" : $value_txt="101~300명"; $a=101; $b=300; break;
						case "v6" : $value_txt="301~"; $a=301; $b=1000000000; break;
						default : $value_txt=""; break;
					}
					if($saddq1!='') $saddq1.=" or";
					$saddq1.=" f.employee_total between ".$a." and ".$b;
					$addstring.="&search[mng_flag][".$key."]=".$value;
					$filter_btn.="<button type='button' class='btn btn-secondary btn-sm' data-field='checkbox|empl|".$key."'>".$value_txt." <span>X</span></button>";
					$i++;
				}
			}
			if($saddq1!='') {
				if($addquery=='') $addquery.=" where";
				else $addquery.=" and";
				if($i==1) {
					$addquery.=$saddq1;
				}else {
					$addquery.=" (".$saddq1.")";
				}
				$addjoinquery1=" left join (SELECT t1.fk_company as fk_company, t1.employee_year as employee_year, t1.employee_total as employee_total FROM `TB_EMPLOYEE` as t1, (select fk_company, max(employee_year) as employee_year from `TB_EMPLOYEE` group by fk_company) as t2 where t1.fk_company=t2.fk_company and t1.employee_year=t2.employee_year) f on f.fk_company=c.company_seq";
				$addjoinquery2=", f.employee_total";
			}
		}

		if($_GET['search']['zip']!='') {
			$saddq1="";
			$i=0;
			foreach($_GET['search']['zip'] as $key => $value) {
				
				if($value!='') {
					switch($key) {
						case "v1" : $value_txt="서울"; $a1=""; break;
						case "v2" : $value_txt="경기"; $a1=""; break;
						case "v3" : $value_txt="인천"; $a1=""; break;
						case "v4" : $value_txt="부산"; $a1=""; break;
						case "v5" : $value_txt="대전"; $a1=""; break;
						case "v6" : $value_txt="대구"; $a1=""; break;
						case "v7" : $value_txt="울산"; $a1=""; break;
						case "v8" : $value_txt="세종"; $a1=""; break;
						case "v9" : $value_txt="광주"; $a1=""; break;
						case "v10" : $value_txt="강원"; $a1=""; break;
						case "v11" : $value_txt="충북"; $a1="충청북도"; break;
						case "v12" : $value_txt="충남"; $a1="충청남도"; break;
						case "v13" : $value_txt="경북"; $a1="경상북도"; break;
						case "v14" : $value_txt="경남"; $a1="경상남도"; break;
						case "v15" : $value_txt="전북"; $a1="전라북도"; break;
						case "v16" : $value_txt="전남"; $a1="전라남도"; break;
						case "v17" : $value_txt="제주"; $a1=""; break;
                        case "v18" : $value_txt="판교"; $a1=""; break;
                        case "v19" : $value_txt="고양"; $a1=""; break;
                        case "v20" : $value_txt="과천"; $a1=""; break;
                        case "v21" : $value_txt="광명"; $a1=""; break;
                        case "v22" : $value_txt="김포"; $a1=""; break;
                        case "v23" : $value_txt="성남"; $a1=""; break;
                        case "v24" : $value_txt="수원"; $a1=""; break;
                        case "v25" : $value_txt="용인"; $a1=""; break;
						default : $value_txt=""; $a1=""; break;
					}
					if($saddq1!='') $saddq1.=" or";
					$saddq1.=" c.company_addr1 like '%".$value_txt."%'";
					if($a1!='') $saddq1.=" or c.company_addr1 like '%".$a1."%'";
					$addstring.="&search[zip][".$key."]=".$value;
					$filter_btn.="<button type='button' class='btn btn-secondary btn-sm' data-field='checkbox|zip|".$key."'>".$value_txt." <span>X</span></button>";
					$i++;
				}
			}
			if($saddq1!='') {
				if($addquery=='') $addquery.=" where";
				else $addquery.=" and";
				if($i==1) {
					$addquery.=" (".$saddq1.")";
				}else {
					$addquery.=" (".$saddq1.")";
				}					
			}
		}
	}

	if(!$_GET['sortcol']) {
		$sortcol='edit_dt';
	}else {
		$sortcol=$_GET['sortcol'];
		$addstring.="&sortcol=".$sortcol;
	}

	if(!$_GET['sortmethod']) {
		$sortmethod='desc';
	}else {
		$sortmethod=$_GET['sortmethod'];
		$addstring.="&sortmethod=".$sortmethod;
	}

	if($sortcol=="membership") $orderby="membership";
	else if($sortcol=="name") $orderby="c.company_nick_ko";
	else if($sortcol=="ceo") $orderby="d.member_nm";
	else if($sortcol=="tel") $orderby="d.affiliation_tel";
	else if($sortcol=="mng_flag") $orderby="c.mng_flag";
	else if($sortcol=='sales') $orderby="e.finance_sales";
	else $orderby="edit_dt";

	//$headers = array('고유값','국문','국문(약칭)','영문','영문(약칭)','주소','전화번호');

	$excel = new PHPExcel();
	$excel->setActiveSheetIndex(0);

	$excel->setActiveSheetIndex(0)->getStyle( "A3:Z4" )->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setARGB('eeeeee');
	$excel->getDefaultStyle()->getFont()->setSize(11);

	$objSheet = $excel->getActiveSheet();

	$objSheet->setTitle('기본정보');
	$objSheet->getCell('A3')->setValue('고유값');
	$objSheet->mergeCells('A3:A4');

	$objSheet->getCell('B3')->setValue('회사명');
	$objSheet->mergeCells('B3:E3');

	$objSheet->getCell('F3')->setValue('주소');
	$objSheet->mergeCells('F3:F4');

	$objSheet->getCell('G3')->setValue('전화번호');
	$objSheet->mergeCells('G3:G4');

	$objSheet->getCell('H3')->setValue('구분');
	$objSheet->mergeCells('H3:I3');

	$objSheet->getCell('J3')->setValue('업종');
	$objSheet->mergeCells('J3:J4');

	$objSheet->getCell('K3')->setValue('설립일자');
	$objSheet->mergeCells('K3:K4');

	$objSheet->getCell('L3')->setValue('로고');
	$objSheet->mergeCells('L3:L4');

	$objSheet->getCell('M3')->setValue('URL');
	$objSheet->mergeCells('M3:M4');

	$objSheet->getCell('N3')->setValue('사업자등록번호');
	$objSheet->mergeCells('N3:N4');

	$objSheet->getCell('O3')->setValue('대표자');
	$objSheet->mergeCells('O3:U3');

	$objSheet->getCell('V3')->setValue('매출액(단위:백만원)');
	$objSheet->mergeCells('V3:Z3');

	$objSheet->getCell('B4')->setValue('국문');
	$objSheet->getCell('C4')->setValue('국문(약칭)');
	$objSheet->getCell('D4')->setValue('영문');
	$objSheet->getCell('E4')->setValue('영문(약칭)');

	$objSheet->getCell('H4')->setValue('회원사여부');
	$objSheet->getCell('I4')->setValue('상장여부');

	$objSheet->getCell('O4')->setValue('이름');
	$objSheet->getCell('P4')->setValue('휴대폰');
	$objSheet->getCell('Q4')->setValue('전화번호');
	$objSheet->getCell('R4')->setValue('이메일');
	$objSheet->getCell('S4')->setValue('생녕월일');
	$objSheet->getCell('T4')->setValue('출신학교');
	$objSheet->getCell('U4')->setValue('주요이력');

	$dateY=date("Y");
	$dateY--;
	$objSheet->getCell('V4')->setValue($dateY);
	$dateY--;
	$objSheet->getCell('W4')->setValue($dateY);
	$dateY--;
	$objSheet->getCell('X4')->setValue($dateY);
	$dateY--;
	$objSheet->getCell('Y4')->setValue($dateY);
	$dateY--;
	$objSheet->getCell('Z4')->setValue($dateY);

	$query="select c.*, if(c.company_nm_ko is null or c.company_nm_ko='',c.company_nick_ko,c.company_nm_ko) as company_title, if(c.company_membership='01','회원사',if(c.company_membership='02','입주사','비회원사')) as membership, if(c.company_edit_dt is null,c.company_reg_dt,company_edit_dt) as edit_dt, d.member_nm, d.affiliation_tel, e.finance_sales, e.finance_year".$addjoinquery2." from TB_COMPANY c left join (select a.fk_company as fk_company, min(b.member_nm) as member_nm, a.affiliation_tel as affiliation_tel from TB_AFFILIATION a left join TB_MEMBER b on a.fk_member=b.member_seq where a.ceo_chk='Y' group by a.fk_company) d on c.company_seq=d.fk_company left join ".$finance_query1." on c.company_seq=e.fk_company".$addjoinquery1.$addquery." order by ".$orderby." ".$sortmethod;
	$result=mysqli_query($db,$query);
	$i=5;
	while($row=mysqli_fetch_array($result)) {
		$common_cd_01_txt="비상장";
		if($row['fk_common_cd_01']!='0' && $row['fk_common_cd_01']!='5') {
			$query1="select*from TB_COMMON_CD where common_cd_seq='".$row['fk_common_cd_01']."'";
			$result1=mysqli_query($db,$query1);
			$row1=mysqli_fetch_array($result1);
			mysqli_free_result($result1);
			$common_cd_01_txt=$row1['common_cd'];
		}
		$common_cd_02_txt="";
		if($row['fk_common_cd_02']!='') {
			$tmp_common_cd_02=explode(",",str_replace("]","",str_replace("[","",$row['fk_common_cd_02'])));
			$tmp_common_cd_02_size=sizeof($tmp_common_cd_02);
			for($j=0;$j<$tmp_common_cd_02_size;$j++) {
				$query1="select*from TB_COMMON_CD where common_cd_seq='".$tmp_common_cd_02[$j]."'";
				$result1=mysqli_query($db,$query1);
				$row1=mysqli_fetch_array($result1);
				mysqli_free_result($result1);
				if($common_cd_02_txt!='') $common_cd_02_txt.=", ";
				$common_cd_02_txt.=$row1['common_cd'];
			}
		}
		if($row['company_setup_dt'] == null || $row['company_setup_dt']=='0000-00-00') {
			$company_setup_dt="-";
		}else {
			$company_setup_dt=$row['company_setup_dt'];
		}

		if($row['company_logo_img_real']!='') {
			$logo_img=$row['company_logo_img_real'];
		}else {
			$logo_img="";
		}
		$img_path="./upload/logo/".$logo_img;

		if($row['company_url'] == null || $row['company_url']=='') {
			$company_url="";
		}else {
			$company_url="http://".str_replace("http://","",str_replace("https://","",$row['company_url']));
		}

		$ceo_info=ceo_info($row['company_seq']);

		if($logo_img!='') { $objSheet->getRowDimension($i)->setRowHeight(35); }
		$objSheet->SetCellValue('A'.$i,$row['company_seq']);
		$objSheet->SetCellValue('B'.$i,$row['company_nm_ko']);
		$objSheet->SetCellValue('C'.$i,$row['company_nick_ko']);
		$objSheet->SetCellValue('D'.$i,$row['company_nm_en']);
		$objSheet->SetCellValue('E'.$i,$row['company_nick_en']);
		$objSheet->SetCellValue('F'.$i,$row['company_addr1']." ".$row['company_addr2']);
		$objSheet->SetCellValue('G'.$i,tel_change($row['company_tel']));
		$objSheet->SetCellValue('H'.$i,$row['membership']);
		$objSheet->SetCellValue('I'.$i,$common_cd_01_txt);
		$objSheet->SetCellValue('J'.$i,$common_cd_02_txt);
		$objSheet->SetCellValue('K'.$i,$company_setup_dt);
		
		if($logo_img!='' && file_exists($img_path)) {
			$objDrawing = new PHPExcel_Worksheet_Drawing();
			$objDrawing->setName('company logo');
			$objDrawing->setDescription('company logo');
			$logo=$img_path;
			$objDrawing->setPath($logo);
			$objDrawing->setOffsetX(10);
			$objDrawing->setOffsetY(10);
			$objDrawing->setCoordinates('L'.$i);
			$objDrawing->setWidth(32);
			$objDrawing->setHeight(32);
			$objDrawing->setWorksheet($excel->getActiveSheet());
		}else {
			$objSheet->SetCellValue('L'.$i,'');
		}
		if($company_url!='') {
			$objSheet->setCellValue('M'.$i, $company_url);
			$objSheet->getCell('M'.$i)->setDataType(PHPExcel_Cell_DataType::TYPE_STRING2);
			$objSheet->getCell('M'.$i)->getHyperlink()->setUrl(strip_tags($company_url));
		}else {
			$objSheet->SetCellValue('M'.$i,'');
		}
		
		$objSheet->SetCellValue('N'.$i,$row['company_reg_no']);

		if (sizeof($ceo_info['member_nm'])>0) {$member_nm=implode( ' / ',$ceo_info['member_nm']);} else {$member_nm='';} 
		if (sizeof($ceo_info['member_hp'])>0) {$member_hp=implode( ' / ',$ceo_info['member_hp']);} else {$member_hp='';} 
		if (sizeof($ceo_info['affiliation_tel'])>0) {$affiliation_tel=implode( ' / ',$ceo_info['affiliation_tel']);} else {$affiliation_tel='';} 
		if (sizeof($ceo_info['member_birth'])>0) {$member_birth=implode( ' / ',$ceo_info['member_birth']);} else {$member_birth='';} 
		if (sizeof($ceo_info['ceo_info_alma_mater'])>0) {$ceo_info_alma_mater=implode( ' / ',$ceo_info['ceo_info_alma_mater']);} else {$ceo_info_alma_mater='';} 
		if (sizeof($ceo_info['ceo_info_history'])>0) {$ceo_info_history=implode( ' / ',$ceo_info['ceo_info_history']);} else {$ceo_info_history='';} 
		if (sizeof($ceo_info['member_id'])>0) {$member_id=implode( ' / ',$ceo_info['member_id']);} else {$member_id='';} 
		
		$objSheet->SetCellValue('O'.$i,$member_nm);
		$objSheet->SetCellValue('P'.$i,$member_hp);
		$objSheet->SetCellValue('Q'.$i,$affiliation_tel);
		$objSheet->SetCellValue('R'.$i,$member_id);
		$objSheet->SetCellValue('S'.$i,$member_birth);
		$objSheet->SetCellValue('T'.$i,$ceo_info_alma_mater);
		$objSheet->SetCellValue('U'.$i,$ceo_info_history);

		$sales_array1=['V','W','X','Y','Z'];
		$dateY=date("Y");
		$dateY--;
		$sales_array2=array();
		$sales_array2[0]=$dateY;
		$dateY--;
		$sales_array2[1]=$dateY;
		$dateY--;
		$sales_array2[2]=$dateY;
		$dateY--;
		$sales_array2[3]=$dateY;
		$dateY--;
		$sales_array2[4]=$dateY;
		
		for($j=0;$j<5;$j++) {
			$fquery1="select * from TB_FINANCE where fk_company='".$row['company_seq']."' and finance_year='".$sales_array2[$j]."' order by finance_seq desc limit 1";
			$fresult1=mysqli_query($db,$fquery1);
			$fnum1=mysqli_num_rows($fresult1);
			if($fnum1==0) {
				$objSheet->SetCellValue($sales_array1[$j].$i,'N/A');
			}else {
				$frow1=mysqli_fetch_assoc($fresult1);
				$objSheet->SetCellValue($sales_array1[$j].$i,(!is_null($frow1['finance_sales']))?number_format($frow1['finance_sales']):"N/A");
			}
		}
		$i++;
	}
	
	$i--;

	// 헤더 칼럼 가운데 정렬
	$objSheet->getStyle('A3:Z4')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
	$objSheet->getStyle('A5:A'.$i)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
	$objSheet->getStyle('G5:G'.$i)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
	$objSheet->getStyle('A3:Z'.$i)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
	// 칼럼 사이즈 자동 조정
	$objSheet->getColumnDimension('A')->setWidth(10);
	$objSheet->getColumnDimension('B')->setWidth(24);  // 칼럼 크기 직접 지정
	$objSheet->getColumnDimension('C')->setWidth(24);
	$objSheet->getColumnDimension('D')->setWidth(24);
	$objSheet->getColumnDimension('E')->setWidth(24);
	$objSheet->getColumnDimension('F')->setWidth(72);
	$objSheet->getColumnDimension('G')->setWidth(15);
	$objSheet->getColumnDimension('H')->setWidth(15);
	$objSheet->getColumnDimension('I')->setWidth(15);
	$objSheet->getColumnDimension('J')->setWidth(15);
	$objSheet->getColumnDimension('K')->setWidth(15);
	$objSheet->getColumnDimension('L')->setWidth(18);
	$objSheet->getColumnDimension('M')->setWidth(35);
	$objSheet->getColumnDimension('N')->setWidth(18);
	$objSheet->getColumnDimension('O')->setWidth(15);
	$objSheet->getColumnDimension('P')->setWidth(15);
	$objSheet->getColumnDimension('Q')->setWidth(15);
	$objSheet->getColumnDimension('R')->setWidth(35);
	$objSheet->getColumnDimension('S')->setWidth(18);
	$objSheet->getColumnDimension('T')->setWidth(25);
	$objSheet->getColumnDimension('U')->setWidth(35);
	$objSheet->getColumnDimension('V')->setWidth(12);
	$objSheet->getColumnDimension('W')->setWidth(12);
	$objSheet->getColumnDimension('X')->setWidth(12);
	$objSheet->getColumnDimension('Y')->setWidth(12);
	$objSheet->getColumnDimension('Z')->setWidth(12);

	$styleArray=array(
		'borders' => array(
			'allborders' => array( 
				'style' => PHPExcel_Style_Border::BORDER_THIN
			)
		)
	);
	$objSheet->getStyle('A3:Z'.$i)->applyFromArray($styleArray);;
	
	//var_dump($excel);
	//exit;

	// 파일 PC로 다운로드
	header('Content-Type: application/octet-stream');
	header('Content-Disposition: attachment; filename="'.$filename.'"');
	header('Cache-Control: max-age=0');

	$objWriter = PHPExcel_IOFactory::createWriter($excel, "Excel2007");
	$objWriter->save('php://output');
	exit;

}else if($_GET['mode']=='hr_view') {
	$query="select*from TB_RESUME where RESUME_SEQ='".$_GET['rseq']."'";
	$result=mysqli_query($db,$query);
	$row=mysqli_fetch_assoc($result);
	$birth=explode(",",birth_txt($row['USER_BIRTH'],'view'));


	$filename=date("YmdHis")."-".$row['USER_NAME']."님의 이력서.xlsx";
	$autoHeight=array();
	$autoHeight2=array();
	$excel = new PHPExcel();
	$excel->setActiveSheetIndex(0);
	$excel->getActiveSheet()->getDefaultRowDimension()->setRowHeight(-1);

	$objSheet = $excel->getActiveSheet();

	$objSheet->setTitle($row['USER_NAME'].'님의 이력서');

	if($row['tmp_img_name']!='') {
		$pic_img=$row['tmp_img_name'];
	}else {
		$pic_img="";
	}
	$img_path="./upload/MEMBER/".$pic_img;

	$i=0;
	if($pic_img!='' && file_exists($img_path)) {
		$objDrawing = new PHPExcel_Worksheet_Drawing();
		$objDrawing->setName('picture');
		$objDrawing->setDescription('picture');
		$pic=$img_path;
		$objDrawing->setPath($pic);
		$objDrawing->setOffsetX(10);
		$objDrawing->setOffsetY(10);
		$objDrawing->setCoordinates('A1');
		//$objDrawing->setWidth(124);
		$objDrawing->setHeight(180);
		$objDrawing->setWorksheet($excel->getActiveSheet());
	}else {
		$objSheet->SetCellValue('A1','');
	}
	$objSheet->mergeCells('A1:D6');

	$objSheet->getCell('E1')->setValue('성 명');
	$objSheet->mergeCells('E1:F2');

	$objSheet->getCell('G1')->setValue('(한글) '.$row['USER_NAME']);
	$objSheet->mergeCells('G1:Q1');

	$objSheet->getCell('G2')->setValue('(영문) '.$row['USER_NAME_EN']);
	$objSheet->mergeCells('G2:Q2');

	$objSheet->getCell('E3')->setValue('생년월일');
	$objSheet->mergeCells('E3:F3');

	$objSheet->getCell('G3')->setValue($birth[0]);
	$objSheet->mergeCells('G3:L3');

	$objSheet->getCell('M3')->setValue('성 별');
	$objSheet->mergeCells('M3:N3');

	$objSheet->getCell('O3')->setValue(($row['USER_SEX']=='M')?"남":"여");
	$objSheet->mergeCells('O3:Q3');

	$objSheet->getCell('E4')->setValue('나 이');
	$objSheet->mergeCells('E4:F4');

	$objSheet->getCell('G4')->setValue(trim($birth[1]));
	$objSheet->mergeCells('G4:L4');

	$objSheet->getCell('M4')->setValue('휴대폰');
	$objSheet->mergeCells('M4:N4');

	$objSheet->getCell('O4')->setValue($row['USER_HP']);
	$objSheet->mergeCells('O4:Q4');

	$objSheet->getCell('E5')->setValue('Email');
	$objSheet->mergeCells('E5:F5');

	$objSheet->getCell('G5')->setValue($row['USER_EMAIL']);
	$objSheet->mergeCells('G5:Q5');

	$objSheet->getCell('E6')->setValue('주 소');
	$objSheet->mergeCells('E6:F6');

	$objSheet->getCell('G6')->setValue($row['USER_ADDR1']." ".$row['USER_ADDR2']);
	$objSheet->mergeCells('G6:Q6');
	$i=6;

	//병역 /보훈 / 장애 start
	$i++;
	$starti=$i;
	$objSheet->getCell('B'.$i)->setValue('역종/군별');
	$objSheet->mergeCells('B'.$i.':C'.$i);
	$objSheet->getStyle('B'.$i)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
	$objSheet->getStyle('B'.$i)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

	$objSheet->getCell('D'.$i)->setValue('복무기간');
	$objSheet->mergeCells('D'.$i.':H'.$i);
	$objSheet->getStyle('D'.$i)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
	$objSheet->getStyle('D'.$i)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

	$objSheet->getCell('I'.$i)->setValue('면제사유');
	$objSheet->mergeCells('I'.$i.':J'.$i);
	$objSheet->getStyle('I'.$i)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
	$objSheet->getStyle('I'.$i)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

	$objSheet->getCell('K'.$i)->setValue('보훈대상');
	$objSheet->mergeCells('K'.$i.':O'.$i);
	$objSheet->getStyle('K'.$i)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
	$objSheet->getStyle('K'.$i)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

	$objSheet->getCell('P'.$i)->setValue('장애대상');
	$objSheet->mergeCells('P'.$i.':Q'.$i);
	$objSheet->getStyle('P'.$i)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
	$objSheet->getStyle('P'.$i)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

	$objSheet->getStyle( "B".$starti.":Q".$starti )->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setARGB('ffffdf');

	$i++;
	array_push($autoHeight,$i);
	$military_period="";
	if($row['MILITARY']=="Y") {
		$military_tmp1=explode("-",$row['MILITARY_SDATE']);
		$military_tmp2=explode("-",$row['MILITARY_EDATE']);
		$military_period=$military_tmp1[0].".".$military_tmp1[1]." ~ ".$military_tmp2[0].".".$military_tmp2[1];
	}
	switch($row['MILITARY']) {
		case "Y" : $tmp_military=$row['MILITARY_TYPE']; $tmp_military_period=""; $tmp_military_rank="(".$row['MILITARY_RANK'].")"; $tmp_exemption=""; break;
		case "N" : $tmp_military="미필"; $tmp_military_period=""; $tmp_military_rank=""; $tmp_exemption=""; break;
		case "E" : $tmp_military="면제"; $tmp_military_period=""; $tmp_military_rank=""; $tmp_exemption=stripslashes($row['EXEMPTION']); break;
		case "X" : $tmp_military="비대상"; $tmp_military_period=""; $tmp_military_rank=""; $tmp_exemption=""; break;
		default : $tmp_military=""; $tmp_military_period=""; $tmp_military_rank=""; $tmp_exemption=""; break;
	}

	$objSheet->getCell('B'.$i)->setValue($tmp_military.$tmp_military_rank);
	$objSheet->mergeCells('B'.$i.':C'.$i);
	$objSheet->getStyle('B'.$i)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
	$objSheet->getStyle('B'.$i)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

	$objSheet->getCell('D'.$i)->setValue($military_period);
	$objSheet->mergeCells('D'.$i.':H'.$i);
	$objSheet->getStyle('D'.$i)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
	$objSheet->getStyle('D'.$i)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

	$objSheet->getStyle('I'.$i)->getAlignment()->setWrapText(true);
	$objSheet->getCell('I'.$i)->setValue($tmp_exemption);
	$objSheet->mergeCells('I'.$i.':J'.$i);
	$objSheet->getStyle('I'.$i)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
	$objSheet->getStyle('I'.$i)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

	$objSheet->getStyle('K'.$i)->getAlignment()->setWrapText(true);
	$objSheet->getCell('K'.$i)->setValue('보훈번호 : '.$row['VETERAN_NUM'].chr(10).'관계 : '.$row['VETERAN_TARGET']);
	$objSheet->mergeCells('K'.$i.':O'.$i);
	$objSheet->getStyle('K'.$i)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
	$objSheet->getStyle('K'.$i)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

	$objSheet->getStyle('P'.$i)->getAlignment()->setWrapText(true);
	$objSheet->getCell('P'.$i)->setValue('장애급수 : '.$row['DISABILITY_GRADE']);
	$objSheet->mergeCells('P'.$i.':Q'.$i);
	$objSheet->getStyle('P'.$i)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
	$objSheet->getStyle('P'.$i)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

	$objSheet->getStyle('A'.$starti)->getAlignment()->setWrapText(true);
	$objSheet->getCell('A'.$starti)->setValue('병'.chr(10).'역');
	$objSheet->mergeCells('A'.$starti.':A'.$i);
	//병역 /보훈 / 장애 end

	// 학력 start
	$i++;
	$starti=$i;
	$objSheet->getCell('B'.$i)->setValue('구분');
	$objSheet->getStyle('B'.$i)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
	$objSheet->getStyle('B'.$i)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

	$objSheet->getCell('C'.$i)->setValue('입학 / 졸업 년월');
	$objSheet->mergeCells('C'.$i.':F'.$i);
	$objSheet->getStyle('C'.$i)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
	$objSheet->getStyle('C'.$i)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);


	$objSheet->getCell('G'.$i)->setValue('출신학교 및 전공');
	$objSheet->mergeCells('G'.$i.':L'.$i);
	$objSheet->getStyle('G'.$i)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
	$objSheet->getStyle('G'.$i)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);


	$objSheet->getCell('M'.$i)->setValue('졸업여부');
	$objSheet->mergeCells('M'.$i.':O'.$i);
	$objSheet->getStyle('M'.$i)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
	$objSheet->getStyle('M'.$i)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);


	$objSheet->getCell('P'.$i)->setValue('소재지');
	$objSheet->getStyle('P'.$i)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
	$objSheet->getStyle('P'.$i)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);


	$objSheet->getCell('Q'.$i)->setValue('성적');
	$objSheet->getStyle('Q'.$i)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
	$objSheet->getStyle('Q'.$i)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);




	$query1="select*from TB_RESUME_EDU where RESUME_SEQ='".$_GET['rseq']."' order by SCH_DIVISION";
	$result1=mysqli_query($db,$query1);
	while($row1=mysqli_fetch_assoc($result1)) {
		$i++;
		switch($row1['SCH_DIVISION']) {
			case "01" : $gubun_txt="고등학교"; break;
			case "02" : $gubun_txt="대학"; break;
			case "03" : $gubun_txt="대학교"; break;
			case "04" : $gubun_txt="대학원"; break;
			case "05" : $gubun_txt="대학원"; break;
			default : $gubun_txt=""; break;
		}

		$pdate=str_replace("-",".",$row1['SCH_INDATE'])." ~ ".$row1['SCH_GRADE_Y'].".".$row1['SCH_GRADE_M'];
		$credit=$row1['SCH_CREDIT1']."/".$row1['SCH_CREDIT2'];

		$objSheet->getCell('B'.$i)->setValue($gubun_txt);
		$objSheet->getStyle('B'.$i)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
		$objSheet->getStyle('B'.$i)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

		$objSheet->getCell('C'.$i)->setValue($pdate);
		$objSheet->mergeCells('C'.$i.':F'.$i);
		$objSheet->getStyle('C'.$i)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
		$objSheet->getStyle('C'.$i)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

		$objSheet->getCell('G'.$i)->setValue($row1['SCH_NAME']);
		$objSheet->mergeCells('G'.$i.':L'.$i);
		$objSheet->getStyle('G'.$i)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
		$objSheet->getStyle('G'.$i)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

		$objSheet->getCell('M'.$i)->setValue($row1['SCH_GRADUATION_STATE']);
		$objSheet->mergeCells('M'.$i.':O'.$i);
		$objSheet->getStyle('M'.$i)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
		$objSheet->getStyle('M'.$i)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

		$objSheet->getCell('P'.$i)->setValue($row1['SCH_LOCATION']);
		$objSheet->getStyle('P'.$i)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
		$objSheet->getStyle('P'.$i)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

		$objSheet->getCell('Q'.$i)->setValue($credit);
		$objSheet->getStyle('Q'.$i)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
		$objSheet->getStyle('Q'.$i)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

	}
	$objSheet->getStyle('A'.$starti)->getAlignment()->setWrapText(true);
	$objSheet->getCell('A'.$starti)->setValue('학'.chr(10).'력');
	$objSheet->mergeCells('A'.$starti.':A'.$i);
	$excel->setActiveSheetIndex(0)->getStyle( "B".$starti.":B".$i )->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setARGB('ffffdf');
	$excel->setActiveSheetIndex(0)->getStyle( "C".$starti.":Q".$starti )->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setARGB('ffffdf');
	$objSheet->getStyle('A'.$starti)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
	$objSheet->getStyle('A'.$starti)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
	// 학력 end

	//이수과목 start
	$temp1=array();
	$_tmp2 = explode("]", str_replace("[", "", $row['SUB_COMPLET'] ) );
	$_tmp2_size=sizeof($_tmp2);
	//echo $_tmp2_size;
	if($row['SUB_COMPLET']!='') {
		for($k=0;$k<$_tmp2_size;$k++) {
			$_tmp3=explode(":",$_tmp2[$k]);
			if(!in_array($_tmp3[0],$temp1)) $temp1[]=$_tmp3[0];
			if(!${"temp_".$_tmp3[0]}) ${"temp_".$_tmp3[0]}=array();
			${"temp_".$_tmp3[0]}[]=$_tmp3[1];
		}
		for($l=0;$l<sizeof($temp1)-1;$l++) {
			$i++;
			if($l==0) $starti=$i;
			$query1="select*from TB_CLASSCODE where seq='".$temp1[$l]."'";
			$result1=mysqli_query($db,$query1);
			$row1=mysqli_fetch_assoc($result1);

			$size=sizeof(${"temp_".$temp1[$l]});
			$_txt="";
			for($m=0;$m<$size;$m++) {
				$query2="select*from TB_CLASSCODE where seq='".${"temp_".$temp1[$l]}[$m]."'";
				$result2=mysqli_query($db,$query2);
				$row2=mysqli_fetch_assoc($result2);
				if($_txt!='') $_txt.=", ";
				$_txt .=$row2['class_name'];
			}

			$objSheet->getCell('B'.$i)->setValue($row1['class_name']);
			$objSheet->mergeCells('B'.$i.':E'.$i);
			$objSheet->getStyle('B'.$i)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

			$objSheet->getCell('F'.$i)->setValue($_txt);
			$objSheet->mergeCells('F'.$i.':Q'.$i);
			$objSheet->getStyle('F'.$i)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
		}
	}
	$excel->getActiveSheet()->getStyle('A'.$starti)->getAlignment()->setWrapText(true);
	$objSheet->getCell('A'.$starti)->setValue('이수'.chr(10).'과목');
	$objSheet->mergeCells('A'.$starti.':A'.$i);
	$objSheet->getStyle('A'.$starti)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
	$objSheet->getStyle('A'.$starti)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
	//이수과목 end

	//프로그램 / 툴 설계경험 start
	$query1="select*from TB_RESUME_PROGRAM where RESUME_SEQ='".$_GET['rseq']."' order by PROGRAM_SEQ";
	$result1=mysqli_query($db,$query1);
	$num1=mysqli_num_rows($result1);
	if($num1 > 0) {
		$i++;
		$starti=$i;
		$excel->setActiveSheetIndex(0)->getStyle( "B".$i.":Q".$i)->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setARGB('ffffdf');
		$objSheet->getCell('B'.$i)->setValue('프로그램 / 툴 명');
		$objSheet->mergeCells('B'.$i.':O'.$i);
		$objSheet->getStyle('B'.$i)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
		$objSheet->getStyle('B'.$i)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

		$objSheet->getCell('P'.$i)->setValue('활용능력');
		$objSheet->mergeCells('P'.$i.':Q'.$i);
		$objSheet->getStyle('P'.$i)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
		$objSheet->getStyle('P'.$i)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

		while($row1=mysqli_fetch_assoc($result1)) {
			$i++;
			switch($row1['PROGRAM_ABILITY']) {
				case "1" : $prog_txt="상"; break;
				case "2" : $prog_txt="중"; break;
				case "3" : $prog_txt="하"; break;
				default : $prog_txt=""; break;
			}

			$objSheet->getCell('B'.$i)->setValue($row1['PROGRAM_NAME']);
			$objSheet->mergeCells('B'.$i.':O'.$i);
			$objSheet->getStyle('B'.$i)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
			$objSheet->getStyle('B'.$i)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

			$objSheet->getCell('P'.$i)->setValue($prog_txt);
			$objSheet->mergeCells('P'.$i.':Q'.$i);
			$objSheet->getStyle('P'.$i)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
			$objSheet->getStyle('P'.$i)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
		}
		$excel->getActiveSheet()->getStyle('A'.$starti)->getAlignment()->setWrapText(true);
		$objSheet->getCell('A'.$starti)->setValue('프로'.chr(10).'그램'.chr(10).'/'.chr(10).'툴'.chr(10).'설계'.chr(10).'경험');
		$objSheet->mergeCells('A'.$starti.':A'.$i);
		$objSheet->getStyle('A'.$starti)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
		$objSheet->getStyle('A'.$starti)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
	}

	// 프로그램 / 툴 설계경험 end

	//대외활동 start
	$query1="select*from TB_RESUME_ACTIVE where RESUME_SEQ='".$_GET['rseq']."' order by CAR_SEQ";
	$result1=mysqli_query($db,$query1);
	$num1=mysqli_num_rows($result1);
	if($num1 > 0) {
		$i++;
		$starti=$i;
		$excel->setActiveSheetIndex(0)->getStyle( "B".$i.":Q".$i)->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setARGB('ffffdf');
		$objSheet->getCell('B'.$i)->setValue('구분');
		$objSheet->mergeCells('B'.$i.':C'.$i);
		$objSheet->getStyle('B'.$i)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
		$objSheet->getStyle('B'.$i)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

		$objSheet->getCell('D'.$i)->setValue('활동기간');
		$objSheet->mergeCells('D'.$i.':F'.$i);
		$objSheet->getStyle('D'.$i)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
		$objSheet->getStyle('D'.$i)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

		$objSheet->getCell('G'.$i)->setValue('주관기관');
		$objSheet->mergeCells('G'.$i.':J'.$i);
		$objSheet->getStyle('G'.$i)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
		$objSheet->getStyle('G'.$i)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

		$objSheet->getCell('K'.$i)->setValue('활동/교육내용');
		$objSheet->mergeCells('K'.$i.':Q'.$i);
		$objSheet->getStyle('K'.$i)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
		$objSheet->getStyle('K'.$i)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

		while($row1=mysqli_fetch_assoc($result1)) {
			$i++;
			array_push($autoHeight,$i);
			$pdate=str_replace("-",".",substr($row1['CAR_SDATE'],0,7))." ~ ".str_replace("-",".",substr($row1['CAR_EDATE'],0,7));

			$objSheet->getCell('B'.$i)->setValue($row1['CAR_CATE']);
			$objSheet->mergeCells('B'.$i.':C'.$i);
			$objSheet->getStyle('B'.$i)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
			$objSheet->getStyle('B'.$i)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

			$objSheet->getCell('D'.$i)->setValue($pdate);
			$objSheet->mergeCells('D'.$i.':F'.$i);			
			$objSheet->getStyle('D'.$i)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
			$objSheet->getStyle('D'.$i)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

			$objSheet->getCell('G'.$i)->setValue($row1['CAR_HOST']);
			$objSheet->mergeCells('G'.$i.':J'.$i);
			$objSheet->getStyle("G".$i)->getAlignment()->setWrapText(true);	
			$objSheet->getStyle('G'.$i)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
			$objSheet->getStyle('G'.$i)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

			$objSheet->getCell('K'.$i)->setValue($row1['CAR_CONTENT']);
			$objSheet->mergeCells('K'.$i.':Q'.$i);
			$objSheet->getStyle("K".$i)->getAlignment()->setWrapText(true);	
			
			$objSheet->getStyle('K'.$i)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
			$objSheet->getStyle('K'.$i)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

		}
		$excel->getActiveSheet()->getStyle('A'.$starti)->getAlignment()->setWrapText(true);
		$objSheet->getCell('A'.$starti)->setValue('대외'.chr(10).'활동');
		$objSheet->mergeCells('A'.$starti.':A'.$i);
		$objSheet->getStyle('A'.$starti)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
		$objSheet->getStyle('A'.$starti)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
	}

	// 대외활동 end

	//수상내역/공모전 start
	$query1="select*from TB_RESUME_AWARD where RESUME_SEQ='".$_GET['rseq']."' order by AWD_SEQ";
	$result1=mysqli_query($db,$query1);
	$num1=mysqli_num_rows($result1);
	if($num1 > 0) {
		$i++;
		$starti=$i;
		$excel->setActiveSheetIndex(0)->getStyle( "B".$i.":Q".$i)->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setARGB('ffffdf');
		
		$objSheet->getCell('B'.$i)->setValue('주최기관');
		$objSheet->mergeCells('B'.$i.':F'.$i);
		$objSheet->getStyle('B'.$i)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
		$objSheet->getStyle('B'.$i)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

		$objSheet->getCell('G'.$i)->setValue('수상일');
		$objSheet->mergeCells('G'.$i.':J'.$i);
		$objSheet->getStyle('G'.$i)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
		$objSheet->getStyle('G'.$i)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

		$objSheet->getCell('K'.$i)->setValue('수상(공모)이름');
		$objSheet->mergeCells('K'.$i.':Q'.$i);
		$objSheet->getStyle('K'.$i)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
		$objSheet->getStyle('K'.$i)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

		while($row1=mysqli_fetch_assoc($result1)) {
			$i++;
			array_push($autoHeight,$i);
			$pdate=str_replace("-",".",$row1['AWD_GETDATE']);

			$objSheet->getCell('B'.$i)->setValue($row1['AWD_HOST']);
			$objSheet->mergeCells('B'.$i.':F'.$i);			
			$objSheet->getStyle('B'.$i)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
			$objSheet->getStyle('B'.$i)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

			$objSheet->getCell('G'.$i)->setValue($pdate);
			$objSheet->mergeCells('G'.$i.':J'.$i);
			$objSheet->getStyle("G".$i)->getAlignment()->setWrapText(true);	
			$objSheet->getStyle('G'.$i)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
			$objSheet->getStyle('G'.$i)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

			$objSheet->getCell('K'.$i)->setValue($row1['AWD_NAME']);
			$objSheet->mergeCells('K'.$i.':Q'.$i);
			$objSheet->getStyle("K".$i)->getAlignment()->setWrapText(true);	
			
			$objSheet->getStyle('K'.$i)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
			$objSheet->getStyle('K'.$i)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

		}
		$excel->getActiveSheet()->getStyle('A'.$starti)->getAlignment()->setWrapText(true);
		$objSheet->getCell('A'.$starti)->setValue('수상'.chr(10).'내역'.chr(10).'/'.chr(10).'공모전');
		$objSheet->mergeCells('A'.$starti.':A'.$i);
		$objSheet->getStyle('A'.$starti)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
		$objSheet->getStyle('A'.$starti)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
	}

	// 수상내역/공모전 end

	//어학 start
	$query1="select*from TB_RESUME_LANG where RESUME_SEQ='".$_GET['rseq']."' order by LAN_SEQ";
	$result1=mysqli_query($db,$query1);
	$num1=mysqli_num_rows($result1);
	if($num1 > 0) {
		$i++;
		$starti=$i;
		$excel->setActiveSheetIndex(0)->getStyle( "B".$i.":Q".$i)->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setARGB('ffffdf');
		$objSheet->getCell('B'.$i)->setValue('구분');
		$objSheet->mergeCells('B'.$i.':C'.$i);
		$objSheet->getStyle('B'.$i)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
		$objSheet->getStyle('B'.$i)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

		$objSheet->getCell('D'.$i)->setValue('외국어명');
		$objSheet->mergeCells('D'.$i.':E'.$i);
		$objSheet->getStyle('D'.$i)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
		$objSheet->getStyle('D'.$i)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

		$objSheet->getCell('F'.$i)->setValue('공인시험');
		$objSheet->mergeCells('F'.$i.':L'.$i);
		$objSheet->getStyle('F'.$i)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
		$objSheet->getStyle('F'.$i)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

		$objSheet->getCell('M'.$i)->setValue('점수/급');
		$objSheet->mergeCells('M'.$i.':N'.$i);
		$objSheet->getStyle('M'.$i)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
		$objSheet->getStyle('M'.$i)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

		$objSheet->getCell('O'.$i)->setValue('취득일');
		$objSheet->mergeCells('O'.$i.':Q'.$i);
		$objSheet->getStyle('O'.$i)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
		$objSheet->getStyle('O'.$i)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

		while($row1=mysqli_fetch_assoc($result1)) {
			$i++;
			switch($row1['LAN_CATEGORY']) {
				case "en" : $lan_category_name="영어"; break;
				case "ja" : $lan_category_name="일본어"; break;
				case "zh" : $lan_category_name="중국어"; break;
				case "de" : $lan_category_name="독일어"; break;
				case "fr" : $lan_category_name="프랑스어"; break;
				case "es" : $lan_category_name="스페인어"; break;
				case "ru" : $lan_category_name="러시아어"; break;
				case "it" : $lan_category_name="이탈리아어"; break;
				case "ar" : $lan_category_name="아랍어"; break;
				case "th" : $lan_category_name="태국어"; break;
				case "ms" : $lan_category_name="마인어"; break;
				case "el" : $lan_category_name="그리스어"; break;
				case "pt" : $lan_category_name="포르투갈어"; break;
				case "vi" : $lan_category_name="베트남어"; break;
				case "nl" : $lan_category_name="네덜란드어"; break;
				case "hi" : $lan_category_name="힌디어"; break;
				case "no" : $lan_category_name="노르웨이어"; break;
				case "hc" : $lan_category_name="유고어"; break;
				case "he" : $lan_category_name="히브리어"; break;
				case "fa" : $lan_category_name="이란(페르시아어)"; break;
				case "tr" : $lan_category_name="터키어"; break;
				case "cs" : $lan_category_name="체코어"; break;
				case "ro" : $lan_category_name="루마니아어"; break;
				case "mn" : $lan_category_name="몽골어"; break;
				case "sv" : $lan_category_name="스웨덴어"; break;
				case "hu" : $lan_category_name="헝가리어"; break;
				case "pl" : $lan_category_name="폴란드어"; break;
				case "my" : $lan_category_name="미얀마어"; break;
				case "sk" : $lan_category_name="슬로바키아어"; break;
				case "sr" : $lan_category_name="세르비아어"; break;
				case "ko" : $lan_category_name="한국어"; break;
				case "direct" : $lan_category_name="직접입력"; break;
				default : $lan_category_name=""; break;
			}
			//$objSheet->getRowDimension($i)->setRowHeight(-1);
			$pdate=str_replace("-",".",$row1['LAN_GETDATE']);

			$objSheet->getCell('B'.$i)->setValue($row1['LAN_TYPE']);
			$objSheet->mergeCells('B'.$i.':C'.$i);
			$objSheet->getStyle('B'.$i)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
			$objSheet->getStyle('B'.$i)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

			$objSheet->getCell('D'.$i)->setValue($lan_category_name);
			$objSheet->mergeCells('D'.$i.':E'.$i);
			$objSheet->getStyle('D'.$i)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
			$objSheet->getStyle('D'.$i)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

			$objSheet->getCell('F'.$i)->setValue($row1['LAN_EXAM']);
			$objSheet->mergeCells('F'.$i.':L'.$i);				
			$objSheet->getStyle("F".$i)->getAlignment()->setWrapText(true);
			$objSheet->getStyle('F'.$i)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
			$objSheet->getStyle('F'.$i)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

			$objSheet->getCell('M'.$i)->setValue($row1['LAN_SCORE']);
			$objSheet->mergeCells('M'.$i.':N'.$i);
			$objSheet->getStyle('M'.$i)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
			$objSheet->getStyle('M'.$i)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

			$objSheet->getCell('O'.$i)->setValue(str_replace("-",".",$row1['LAN_GETDATE']));
			$objSheet->mergeCells('O'.$i.':Q'.$i);			
			$objSheet->getStyle('O'.$i)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
			$objSheet->getStyle('O'.$i)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
		}
		$excel->getActiveSheet()->getStyle('A'.$starti)->getAlignment()->setWrapText(true);
		$objSheet->getCell('A'.$starti)->setValue('어'.chr(10).'학');
		$objSheet->mergeCells('A'.$starti.':A'.$i);
		$objSheet->getStyle('A'.$starti)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
		$objSheet->getStyle('A'.$starti)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
	}

	// 어학 end

	//자격증 start
	$query1="select*from TB_RESUME_LICENSE where RESUME_SEQ='".$_GET['rseq']."' order by LIC_SEQ";
	$result1=mysqli_query($db,$query1);
	$num1=mysqli_num_rows($result1);
	if($num1 > 0) {
		$i++;
		$starti=$i;
		$excel->setActiveSheetIndex(0)->getStyle( "B".$i.":Q".$i)->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setARGB('ffffdf');
		$objSheet->getCell('B'.$i)->setValue('발행처');
		$objSheet->mergeCells('B'.$i.':E'.$i);
		$objSheet->getStyle('B'.$i)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
		$objSheet->getStyle('B'.$i)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

		$objSheet->getCell('F'.$i)->setValue('자격증명');
		$objSheet->mergeCells('F'.$i.':L'.$i);
		$objSheet->getStyle('F'.$i)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
		$objSheet->getStyle('F'.$i)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

		$objSheet->getCell('M'.$i)->setValue('점수/급');
		$objSheet->mergeCells('M'.$i.':N'.$i);
		$objSheet->getStyle('M'.$i)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
		$objSheet->getStyle('M'.$i)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

		$objSheet->getCell('O'.$i)->setValue('취득일');
		$objSheet->mergeCells('O'.$i.':Q'.$i);
		$objSheet->getStyle('O'.$i)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
		$objSheet->getStyle('O'.$i)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

		while($row1=mysqli_fetch_assoc($result1)) {
			$i++;
			//$objSheet->getRowDimension($i)->setRowHeight(-1);
			$pdate=str_replace("-",".",$row1['LIC_GETDATE']);

			$objSheet->getCell('B'.$i)->setValue($row1['LIC_HOST']);
			$objSheet->mergeCells('B'.$i.':E'.$i);
			$objSheet->getStyle('B'.$i)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
			$objSheet->getStyle('B'.$i)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

			$objSheet->getCell('F'.$i)->setValue($row1['LIC_NAME']);
			$objSheet->mergeCells('F'.$i.':L'.$i);				
			$objSheet->getStyle("F".$i)->getAlignment()->setWrapText(true);
			$objSheet->getStyle('F'.$i)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
			$objSheet->getStyle('F'.$i)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

			$objSheet->getCell('M'.$i)->setValue($row1['LIC_SCORE']);
			$objSheet->mergeCells('M'.$i.':N'.$i);
			$objSheet->getStyle('M'.$i)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
			$objSheet->getStyle('M'.$i)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

			$objSheet->getCell('O'.$i)->setValue(str_replace("-",".",$row1['LIC_GETDATE']));
			$objSheet->mergeCells('O'.$i.':Q'.$i);			
			$objSheet->getStyle('O'.$i)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
			$objSheet->getStyle('O'.$i)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
		}
		$excel->getActiveSheet()->getStyle('A'.$starti)->getAlignment()->setWrapText(true);
		$objSheet->getCell('A'.$starti)->setValue('자'.chr(10).'격'.chr(10).'증');
		$objSheet->mergeCells('A'.$starti.':A'.$i);
		$objSheet->getStyle('A'.$starti)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
		$objSheet->getStyle('A'.$starti)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
	}

	// 자격증 end

	//경력 start
	$query1="select*from TB_RESUME_CAREER where RESUME_SEQ='".$_GET['rseq']."' order by CAR_SEQ";
	$result1=mysqli_query($db,$query1);
	$num1=mysqli_num_rows($result1);
	if($num1 > 0) {
		$i++;
		$starti=$i;
		$excel->setActiveSheetIndex(0)->getStyle( "B".$i.":Q".$i)->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setARGB('ffffdf');
		$objSheet->getCell('B'.$i)->setValue('직장명');
		$objSheet->mergeCells('B'.$i.':C'.$i);
		$objSheet->getStyle('B'.$i)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
		$objSheet->getStyle('B'.$i)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

		$objSheet->getCell('D'.$i)->setValue('근무기간');
		$objSheet->mergeCells('D'.$i.':F'.$i);
		$objSheet->getStyle('D'.$i)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
		$objSheet->getStyle('D'.$i)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

		$objSheet->getCell('G'.$i)->setValue('부서');
		$objSheet->mergeCells('G'.$i.':J'.$i);
		$objSheet->getStyle('G'.$i)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
		$objSheet->getStyle('G'.$i)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

		$objSheet->getCell('K'.$i)->setValue('직급');
		$objSheet->mergeCells('K'.$i.':L'.$i);
		$objSheet->getStyle('K'.$i)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
		$objSheet->getStyle('K'.$i)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

		$objSheet->getCell('M'.$i)->setValue('담당업무');
		$objSheet->mergeCells('M'.$i.':Q'.$i);
		$objSheet->getStyle('M'.$i)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
		$objSheet->getStyle('M'.$i)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

		while($row1=mysqli_fetch_assoc($result1)) {
			$i++;
			array_push($autoHeight,$i);
			//$objSheet->getRowDimension($i)->setRowHeight(-1);
			$pdate=str_replace("-",".",substr($row1['CAR_SDATE'],0,7))." ~ ".str_replace("-",".",substr($row1['CAR_EDATE'],0,7));

			$objSheet->getCell('B'.$i)->setValue($row1['CAR_HOST']);
			$objSheet->mergeCells('B'.$i.':C'.$i);
			$objSheet->getStyle('B'.$i)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
			$objSheet->getStyle('B'.$i)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

			$objSheet->getCell('D'.$i)->setValue($pdate);
			$objSheet->mergeCells('D'.$i.':F'.$i);
			$objSheet->getStyle('D'.$i)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
			$objSheet->getStyle('D'.$i)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

			$objSheet->getCell('G'.$i)->setValue($row1['CAR_DEPT']);
			$objSheet->mergeCells('G'.$i.':J'.$i);
			$objSheet->getStyle('G'.$i)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
			$objSheet->getStyle('G'.$i)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

			$objSheet->getCell('K'.$i)->setValue($row1['CAR_POSITION']);
			$objSheet->mergeCells('K'.$i.':L'.$i);
			$objSheet->getStyle('K'.$i)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
			$objSheet->getStyle('K'.$i)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

			$objSheet->getCell('M'.$i)->setValue($row1['CAR_DUTY']);
			$objSheet->mergeCells('M'.$i.':Q'.$i);			
			$objSheet->getStyle("M".$i)->getAlignment()->setWrapText(true);			
			
			$objSheet->getStyle('M'.$i)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
			$objSheet->getStyle('M'.$i)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
		}
		$excel->getActiveSheet()->getStyle('A'.$starti)->getAlignment()->setWrapText(true);
		$objSheet->getCell('A'.$starti)->setValue('경'.chr(10).'력');
		$objSheet->mergeCells('A'.$starti.':A'.$i);
		$objSheet->getStyle('A'.$starti)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
		$objSheet->getStyle('A'.$starti)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
	}

	// 경력 end

	//자기소개서 start
	$i++;
	$k=$i+10;
	for($j=$i;$j<=$i+10;$j++) {
		array_push($autoHeight,$j);
	}
	//$objSheet->getRowDimension($i)->setRowHeight(-1);
	$excel->getActiveSheet()->getStyle('A'.$i)->getAlignment()->setWrapText(true);
	$objSheet->getCell('A'.$i)->setValue('자'.chr(10).'기'.chr(10).'소'.chr(10).'개'.chr(10).'서');
	$objSheet->mergeCells('A'.$i.':A'.$k);
	$objSheet->getStyle('A'.$i)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
	$objSheet->getStyle('A'.$i)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

	$objSheet->getCell('B'.$i)->setValue($row['RESUME_PROFILE']);
	$objSheet->mergeCells('B'.$i.':Q'.$k);
	$objSheet->getRowDimension($i)->setRowHeight(-1);	
	$objSheet->getStyle("B".$i)->getAlignment()->setWrapText(true);			
	
	//$objSheet->getStyle('B'.$i)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
	$objSheet->getStyle('B'.$i)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
	//자기소개서 end
	$i=$k;

	$excel->setActiveSheetIndex(0)->getStyle( "E1:F2" )->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setARGB('fff9ee');
	$excel->setActiveSheetIndex(0)->getStyle( "E3:F3" )->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setARGB('fff9ee');
	$excel->setActiveSheetIndex(0)->getStyle( "M3:N3" )->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setARGB('fff9ee');
	$excel->setActiveSheetIndex(0)->getStyle( "E4:F4" )->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setARGB('fff9ee');
	$excel->setActiveSheetIndex(0)->getStyle( "M4:N4" )->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setARGB('fff9ee');
	$excel->setActiveSheetIndex(0)->getStyle( "E5:F5" )->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setARGB('fff9ee');
	$excel->setActiveSheetIndex(0)->getStyle( "E6:F6" )->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setARGB('fff9ee');
	$excel->setActiveSheetIndex(0)->getStyle( "A7:A".$i )->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setARGB('e1ffe3');
	//$excel->setActiveSheetIndex(0)->getStyle( "B7:Q7" )->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setARGB('ffffdf');
	

	//컬럼 가로 사이즈
	$objSheet->getColumnDimension('A')->setWidth(9.83);
	$objSheet->getColumnDimension('B')->setWidth(9.83);
	$objSheet->getColumnDimension('C')->setWidth(3.33);	
	$objSheet->getColumnDimension('D')->setWidth(2);
	$objSheet->getColumnDimension('E')->setWidth(13.17);
	$objSheet->getColumnDimension('F')->setWidth(4.67);
	$objSheet->getColumnDimension('G')->setWidth(4.33);
	$objSheet->getColumnDimension('H')->setWidth(3.83);
	$objSheet->getColumnDimension('I')->setWidth(5.33);
	$objSheet->getColumnDimension('J')->setWidth(9.67);
	$objSheet->getColumnDimension('K')->setWidth(9);
	$objSheet->getColumnDimension('L')->setWidth(0.45);
	$objSheet->getColumnDimension('M')->setWidth(9.17);
	$objSheet->getColumnDimension('N')->setWidth(3.17);
	$objSheet->getColumnDimension('O')->setWidth(3.33);
	$objSheet->getColumnDimension('P')->setWidth(9.17);
	$objSheet->getColumnDimension('Q')->setWidth(9.67);

	//컬럼 세로 사이즈
	$objSheet->getRowDimension('1')->setRowHeight(27);
	$objSheet->getRowDimension('2')->setRowHeight(27);
	$objSheet->getRowDimension('3')->setRowHeight(24);
	$objSheet->getRowDimension('4')->setRowHeight(24);
	$objSheet->getRowDimension('5')->setRowHeight(22);
	$objSheet->getRowDimension('6')->setRowHeight(21);

	//$excel->getActiveSheet()->getDefaultRowDimension()->setRowHeight(-1);
	for($j=7;$j<=$i;$j++) {
		if(!in_array($j,$autoHeight) && !in_array($j,$autoHeight2)) {
			$objSheet->getRowDimension($j)->setRowHeight(21);
		}
	}

	//var_dump($autoHeight);
	//exit;

	for($j=0;$j<sizeof($autoHeight);$j++) {
		$objSheet->getRowDimension($autoHeight[$j])->setRowHeight(40);
	}

	for($j=0;$j<sizeof($autoHeight2);$j++) {
		//$objSheet->getRowDimension($autoHeight2[$j])->setRowHeight(240);
	}
	
	$styleArray=array(
		'borders' => array(
			'allborders' => array( 
				'style' => PHPExcel_Style_Border::BORDER_THIN
			)
		)
	);

	//정렬
	$objSheet->getStyle('A1')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
	$objSheet->getStyle('E1')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
	$objSheet->getStyle('E3')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
	$objSheet->getStyle('M3')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
	$objSheet->getStyle('G3')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
	$objSheet->getStyle('O3')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
	$objSheet->getStyle('E4')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
	$objSheet->getStyle('G4')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
	$objSheet->getStyle('O4')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
	$objSheet->getStyle('M4')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
	$objSheet->getStyle('E5')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
	$objSheet->getStyle('E6')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
	$objSheet->getStyle('A7')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
	$objSheet->getStyle('B7')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
	$objSheet->getStyle('C7')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
	$objSheet->getStyle('G7')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
	$objSheet->getStyle('M7')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
	$objSheet->getStyle('P7')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
	$objSheet->getStyle('Q7')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
	$objSheet->getStyle('E1')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
	$objSheet->getStyle('E3')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
	$objSheet->getStyle('G1')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
	$objSheet->getStyle('G2')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
	$objSheet->getStyle('G3')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
	$objSheet->getStyle('M3')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
	$objSheet->getStyle('O3')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
	$objSheet->getStyle('E4')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
	$objSheet->getStyle('G4')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
	$objSheet->getStyle('M4')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
	$objSheet->getStyle('O4')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
	$objSheet->getStyle('E5')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
	$objSheet->getStyle('G5')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
	$objSheet->getStyle('E6')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
	$objSheet->getStyle('G6')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
	$objSheet->getStyle('A7')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
	$objSheet->getStyle('B7')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
	$objSheet->getStyle('C7')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
	$objSheet->getStyle('G7')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
	$objSheet->getStyle('M7')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
	$objSheet->getStyle('P7')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
	$objSheet->getStyle('Q7')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

	
	$objSheet->getStyle('A1:Q'.$i)->applyFromArray($styleArray);

	header('Content-Type: application/octet-stream');
	header('Content-Disposition: attachment; filename="'.$filename.'"');
	header('Cache-Control: max-age=0');

	$objWriter = PHPExcel_IOFactory::createWriter($excel, "Excel2007");
	$objWriter->save('php://output');
	exit;
}




// 엑셀 생성
$last_char = column_char( count($headers) - 1 );
 
$excel = new PHPExcel();
$excel->setActiveSheetIndex(0)->getStyle( "A1:${last_char}1" )->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setARGB($header_bgcolor);
$excel->setActiveSheetIndex(0)->getStyle( "A:$last_char" )->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)->setWrapText(true);
foreach($widths as $i => $w) $excel->setActiveSheetIndex(0)->getColumnDimension( column_char($i) )->setWidth($w);
$excel->getActiveSheet()->getStyle("A1:${last_char}1")->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$excel->getActiveSheet()->fromArray($data,NULL,'A1');
$styleArray=array(
	'borders' => array(
		'allborders' => array( 
			'style' => PHPExcel_Style_Border::BORDER_THIN
		)
	)
);

$excel->getActiveSheet()->getStyle("A1:${last_char}".$num)->applyFromArray($styleArray);
//$excel->getActiveSheet()->getStyle("C1:C".$num)->getNumberFormat()->setFormatCode("#,###,###");
if($_GET['mode']=="company") {
	$excel->getActiveSheet()->getStyle("B1:B".$num)->getNumberFormat()->setFormatCode("###-##-#####");
}
$writer = PHPExcel_IOFactory::createWriter($excel, 'Excel2007');

header('Content-Type: application/octet-stream');
header('Content-Disposition: attachment; filename="'.$filename.'"');
header('Cache-Control: max-age=0');

$writer->save('php://output');

@mysqli_close($db);
?>