package wordReader.biProject;

import org.bouncycastle.jce.provider.JDKDSASigner.ecDSA;

public class Time {
	
	int hours ; 
	int minutes ;
	
	public Time(int hours, int minutes) {
		super();
		this.hours = hours;
		this.minutes = minutes;
	}
	
	public static Time diffTime( Time start, Time stop) {
		Time diffTime = new Time(0, 0) ;
		
		if( stop.minutes >= start.minutes) {
			diffTime.hours = stop.hours - start.hours ;
			diffTime.minutes = stop.minutes - start.minutes ;
		}
		else {
			diffTime.hours = stop.hours - start.hours - 1 ;
			diffTime.minutes = ( 60 - start.minutes ) + stop.minutes ;
		}
		
		return diffTime ;
	}
	
	public static Time getTotalHourNMins( Time start, Time stop) {
		// 取得總共時間
		Time totalTime = diffTime(start, stop) ;
		
		// 扣掉中午休息分鐘數 
		Time lunchTime = new Time(12, 0);
		Time lunchTime2 = new Time(13, 0);
		Time dinnerTime = new Time(18, 0);
		Time dinnerTime2 = new Time(19, 0);
		Time minusTime = new Time(0, 0) ;
		
		// S E 中午
//		if( start.hours < 12 && stop.hours < 12 ) {
//			// 不扣, 因為沒有經過中午時間
//			return totalTime ;
//		}
		// S  | E 中午|
		if( start.hours < 12 && stop.hours == 12 ) {
			minusTime = diffTime(lunchTime, stop) ; 
		}
		// | s 中午 |  E
		else if( start.hours == 12 && stop.hours >= 13) {
			minusTime = diffTime(start, lunchTime2) ;
//			if( stop.hours < 18 ) { // 沒經過晚上不需考慮
//				return minus(totalTime, minusTime) ;
//			}
			if( stop.hours == 18 ) { // 剛好在晚上區間
				minusTime = Add(minusTime, diffTime(dinnerTime, stop) );
			}
			else if( stop.hours >= 19 ) {
				minusTime = Add(minusTime, diffTime(dinnerTime, dinnerTime2)) ;
			}
		}
		// s |中午| E
		else if( start.hours < 12 && stop.hours >= 13) {
			minusTime = diffTime(lunchTime, lunchTime2) ;
			
//			if( stop.hours < 18 ) { // 沒經過晚上不需考慮
//				return minus(totalTime, minusTime) ;
//			}
			if( stop.hours == 18 ) { // 剛好在晚上區間
				minusTime = Add(minusTime, diffTime(dinnerTime, stop) );
			}
			else if( stop.hours >= 19 ) {
				minusTime = Add(minusTime, diffTime(dinnerTime, dinnerTime2)) ;
			}
		}
		// |中午| S E
		else if( start.hours >= 13 && stop.hours >= 13) {
			// 不扣, 因為沒有經過中午時間
			
			// se|晚上|
//			if(start.hours < 18 && stop.hours < 18) {
//				// 不扣, 因為沒經過晚上時間
//				return totalTime ;
//			}
			// s|晚上|e
			if( start.hours < 18 && stop.hours >= 19 ) {
				minusTime = diffTime(dinnerTime, dinnerTime2);
			}
			// s |晚上 e|
			else if( start.hours < 18 && stop.hours == 18 ) {
				minusTime = diffTime(dinnerTime, stop) ;
			}
			// |晚上s| e
			else if( start.hours == 18 && stop.hours >= 19) {
				minusTime = diffTime(start, dinnerTime2) ;
			}
			// ｜晚上｜se
//			else if( start.hours >= 19 && stop.hours >= 19 ) {
//				// 不扣, 因為沒經過晚上時間
//				return totalTime ;
//			}
		}
		
		return minus(totalTime, minusTime) ;

	}
	
	// return (Time a minus minutes)
	public static Time minus(Time a,  Time b ) {
		
		Time ret_Time = new Time(0, 0);
		
		if( a.hours == 0 ) {
			ret_Time.minutes = a.minutes - b.minutes ;
		}
		else if( a.hours > 0 ) {
			a.hours-- ;//扣掉一小時
			a.minutes = a.minutes + 60 ; // 補到分鐘數上面
			ret_Time.hours = a.hours - b.hours ;
			ret_Time.minutes = a.minutes - b.minutes ;
			// 相減完分鐘之後,判斷分鐘是否大於60
			if(ret_Time.minutes >= 60) {
				ret_Time.minutes = ret_Time.minutes - 60 ;
				ret_Time.hours = ret_Time.hours + 1 ;
			}
		}
	
		return ret_Time ;
	}

	
	public static Time Add( Time a, Time b) {
		
		Time retTime = new Time(0, 0);
		
		retTime.hours = a.hours + b.hours ;
		retTime.minutes = a.minutes + b.minutes ; 
		
		if( retTime.minutes >= 60 ) {
			retTime.minutes = retTime.minutes - 60 ; 
			retTime.hours = retTime.hours + 1 ; 
		}
		
		return retTime ;
	}
}
