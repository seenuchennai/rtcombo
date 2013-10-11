using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FileHelpers;

namespace Shubharealtime
{
    


        [DelimitedRecord(","), IgnoreFirst(7), IgnoreEmptyLines(true)]
        public class GOOGLE
        {
            public string Name;
            [FieldNullValue(typeof(string), "0")]

            [FieldOptional()]

            public string CLOSE_PRICE;

            [FieldNullValue(typeof(string), "0")]
            [FieldOptional()]
            public string HIGH_PRICE;
            [FieldOptional()]
            [FieldNullValue(typeof(string), "0")]

            public string LOW_PRICE;
            [FieldNullValue(typeof(string), "0")]
            [FieldOptional()]
            public string OPEN_PRICE;



            [FieldOptional()]
            [FieldNullValue(typeof(string), "0")]

            public string volume;
            


        }
        [DelimitedRecord(","), IgnoreFirst(7), IgnoreEmptyLines(true)]
        public class nestnow
        {
            public string Name;
            public string datetime;
            [FieldNullValue(typeof(string), "0")]
            [FieldOptional()]
            public string OPEN_PRICE;
            [FieldNullValue(typeof(string), "0")]
            [FieldOptional()]
            public string HIGH_PRICE;
            [FieldOptional()]
            [FieldNullValue(typeof(string), "0")]

            public string LOW_PRICE;
            [FieldNullValue(typeof(string), "0")]

            [FieldOptional()]

            public string CLOSE_PRICE;
                      
                [FieldOptional()]
            [FieldNullValue(typeof(string), "0")]

            public string volume;



        }
        [DelimitedRecord(","), IgnoreFirst(1)]
        public class GOOGLEFINAL
        {
            public string ticker;
            public string name;
            public string date;
            public string time;
           

            public string open;
            public string high;
            public string low;
            public string close;
            public string volume;
            [FieldNullValue(typeof(long), "0")]
            public Nullable<long> openint;
        }
        [DelimitedRecord(","), IgnoreFirst(1)]
        public class nestnowfinal
        {
            public string ticker;
            public string date;
            public string time;


            public string open;
            public string high;
            public string low;
            public string close;
            public string volume;
            
        }
    
    
}
