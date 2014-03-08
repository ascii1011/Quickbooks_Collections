<?
/////////////////////////////////////////////////////////////////////
//  IMPORTANT!!!!!   IMPORTANT!!!!!   IMPORTANT!!!!!!            
/////////////////////////////////////////////////////////////////////
//
//
//  DO NOT REMOVE OR EDIT THIS FILE!!!!!!!!!!!!!!!!!!!!!!!!!!!!!   
//
//  THIS FILE IS AN ITEGRATED PART OF THE INTERNAL COLLECTIONS     
//               VISUAL BASIC APPLICATION.
//
//
/////////////////////////////////////////////////////////////////////
//  This file is for pulling the status of one dealer per request  
//  
//  Input Dealer ID
//	ex. http://test.com/cmc.php?id=5454633
//
//  return Dealer status value
//
//  By: Christopher Harty
/////////////////////////////////////////////////////////////////////

function pullSingleStatus($temp_id)
{
 		 //connect to DB
     $db=mysql_connect("localhost", "root", "ascii1011");
     mysql_select_db("autoloan", $db);
		 
     //Query
  	 $query="SELECT * FROM dealers WHERE ID=".$temp_id;
  	 
  	 //result
  	 $result = mysql_query($query,$db);
  	 $row = mysql_fetch_array($result);
  
  	 //Value returned
  	 $dstatus=$row['DSTATUS'];
		 
		 mysql_close($db);
        
     return $temp_id . ":" . $dstatus;
}

set_time_limit(0);

if ($id) {

	 $result = "";
   $many_ids = preg_split('/-/', $id);
	 
	 if (sizeof($many_ids) > 0) {	 		  	 	 			 
			 for ($i=0;$i<sizeof($many_ids);$i++) {
			 		 //print $many_ids[$i]."<br>";
			 		if ($i+1 == sizeof($many_ids)) {
   		    	 $result .= pullSingleStatus($many_ids[$i]);
					} else {
   		    	 $result .= pullSingleStatus($many_ids[$i])."-";
					}
			 }
	 } else {			 
	 		 $result = pullSingleStatus($id);
	 }
	 	 
	 print "<br>".$result."<br>";
	 
	 //when grabbing value, split by "<br>" and grab array[1]

}

?>