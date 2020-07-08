/*  
Copyright 2019 Ted Tyree - ieSimplified.com

Permission is hereby granted, free of charge, to any person obtaining a copy of this 
software and associated documentation files (the "Software"), to deal in the Software
without restriction, including without limitation the rights to use, copy, modify, 
merge, publish, distribute, sublicense, and/or sell copies of the Software, and to 
permit persons to whom the Software is furnished to do so, subject to the 
following conditions:

The above copyright notice and this permission notice shall be included in all copies 
or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, 
INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A 
PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT 
HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION 
OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE 
OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
*/
// iesJSON version 4 - Simple light-weight JSON Class for .aspx and .ashx web pages
// Copyright 2015 - Ted Tyree - ieSimplified, LLC - All rights reserved.
// ****************************************************
// NameSpace: iesJSONlib

// **************************************************************************
// ***************************  iesJSON  ************************************
// **************************************************************************
//
// This is a new version of iesJSON that makes each 'Node' in the JSON Object/Array a JSON Object with a Key/Value pair.
// So an "object" is a dynamic array of iesJSON objects (key/value pairs)
// ... an "array" is a dynamic array of iesJSON objects (key/value pairs with all key values = null)
// ... a "string" is an iesJSON node with key=null and value=<the string>
// ... a null is an iesJSON node with key=null and value=null
// etc.
//
// You can implement foreach with this class:
//    iesJSON i=new iesJSON("[5,4,3,2,1]");
//    foreach (object k in i) {
//		...
//
// You can reference items using [index]...
//    iesJSON i=new iesJSON("{'color1':'black','color2':'blue','dogs':['Pincher','Sausage','Doverman','Chiwawa']}");
//    MessageBox.Show("First color: " + i["color1"]);
//    MessageBox.Show("Second dog: " + i["dogs"][2]);
//    MessageBox.Show("... and again: " + i[2][2]);
//    MessageBox.Show("... and yet again: " + i["dogs.2"]);
//
// This is in addition to the old method of referencing an item by text...
//    MessageBox.Show("Second dog: " + i.GetString("dogs.2"));
//    MessageBox.Show("... and again : " + i.GetString("3.2"));
//
// Warning!  This version does not naturally 'sort' the object array.
// FUTURE: It might be a good idea to create a KeepSorted flag so that you can cause the JSON object to always be sorted (also making gets/updates more efficient with a binary search)
// FUTURE: Allow comments in the JSON string? (probably similar to c-sharp comments)
// FUTURE: As part of the StatusMessage (or a separate field) show the text where the parse failed to make it easy to debug the JSON code.  Maybe the text just before the error, too.
//

namespace iesJSONlib
{
    using System;
    using System.Collections;
    using System.Text;
    using System.IO;
    //using Newtonsoft.Json;

    public class iesJSON : IEnumerable, IComparable
    {

        // *** if value_valid=false and jsonString_valid=false then this JSON Object is null (not set to anything)
        private int _status = 0;  // *** 0=OK, anything else represents an error/invalid status

        //FUTURE: REMOVE THESE ONCE THE stats OBJECT IS IMPLEMENTED
        private string _statusMsg = "";  // *** Error message will be included in the stats object (if there is an error)
        public string tmpStatusMsg = ""; // DEBUG DEBUG DEBUG

        private string _jsonType = "";  // object, array, string, number, boolean, null, error
        private string _key = null;  // only used for JSON objects.  all other _jsonType values will have a key value of null
        private object _value = null;
        private bool _value_valid = false;
        private string _jsonString = "";
        private bool _jsonString_valid = false;
        public int endpos = 0;
        public iesJSON Parent;
        public bool ALLOW_SINGLE_QUOTE_STRINGS = true;
        public bool ENCODE_SINGLE_QUOTES = false;

        // jsonTypeEnum Values - Numeric representation of jsonType
        //NOTE: The NUMBER VALUES are critical to the SORT() routine
        //  So that iesJSON arrays get sorted in this order: NULL, Number, String, Boolean, Array, Object, Error
        public const int jsonTypeEnum_null = 0;
        public const int jsonTypeEnum_number = 1;
        public const int jsonTypeEnum_string = 2;
        public const int jsonTypeEnum_boolean = 3;
        public const int jsonTypeEnum_array = 10;
        public const int jsonTypeEnum_object = 11;
        public const int jsonTypeEnum_error = 90;
        public const int jsonTypeEnum_invalid = 99;

        // The _keep object is normally NULL indicating that no spacing or comments should be captured or rendered
        // When/if this object gets created, it will include flags to indicate if the object is intended to keep spacing and/or comments
        //   keepSpacing - false=DO NOT preserve spacing, true=Preserve spacing including carriage returns
        //   keepComments - false=DO NOT preseerve comments, true=Preserve comments in Flex JSON
        // NOTE: These flags only have an affect during the deserialize process.
        // NEW: KEEP feature also retains the lineNumber and the linePosition of each iesJSON item during the parse process
        // FUTURE: How to clear spacing/comments after the load?  (set _keep to null? or just keep/pre/post attributes) Make this recursive to sub-objects?
        // FUTURE: Flag to keep double/single quotes or no-quotes for Flex JSON (and write back to file using same quotes or no-quote per item)
        private iesJSON _keep = null;
        private bool _UseFlexJson = false;  //FUTURE: FlexJson should also change the way we serialize.  Right now it only changes the way we deserialize 03/2015

        public iesJSON stats = null;  // null indicates we are not tracking stats.  stats object will also be created if an error occurs - to hold the error message.
        private bool _NoStatsOrMsgs = false;  // This flag is necessary so that the stats do not track stats. (set to True for stats object)
                                              // Below are the stats that will be included in the stats object.
                                              //FUTURE: REMOVE THESE ONCE THE stats OBJECT IS IMPLEMENTED
        /*
            public int stat_clear=0;
            public int stat_clearFromSelf=0;
            public int stat_clearFromOther=0;
            public int stat_serialize=0;
            public int stat_serializeme=0;
            public int stat_Deserialize=0;
            public int stat_DeserializeMe=0;
            public int stat_getobj=0;
            public int stat_invalidate=0;
            public int stat_invalidateFromSelf=0;
            public int stat_invalidateFromChild=0;
            public int stat_invalidateFromOther=0;
        */

        // CONSTRUCTORS - includes ability to initialize the JSON object with a JSON string for example j=new iesJSON("{}");
        public iesJSON() { /* nothign to do here */ }
        public iesJSON(bool UseFlexJsonFlag, string InitialJSON)
        {
            UseFlexJson = UseFlexJsonFlag;
            Deserialize(InitialJSON);
        }
        public iesJSON(string InitialJSON)
        {
            Deserialize(InitialJSON);
        }
        public iesJSON(bool UseFlexJsonFlag)
        {
            UseFlexJson = UseFlexJsonFlag;
        }

        // CreateEmptyObject()
        // This accomplishes the same thing as iesJSON("{}") but avoids recursive loops within the class by simply setting up the Object without any parsing.
        // This is also faster than calling iesJSON("{}")
        static public iesJSON CreateEmptyObject()
        {
            iesJSON j = new iesJSON();
            j._value = new System.Collections.Generic.List<object>();
            j._jsonType = "object";
            j._value_valid = true;
            j._jsonString = "{}";
            j._jsonString_valid = false;
            //j._status already = 0
            return j;
        }

        // CreateEmptyArray()
        // This accomplishes the same thing as iesJSON("[]") but avoids recursive loops within the class by simply setting up the Array without any parsing.
        // This is also faster than calling iesJSON("[]")
        static public iesJSON CreateEmptyArray()
        {
            iesJSON j = new iesJSON();
            j._value = new System.Collections.Generic.List<object>();
            j._jsonType = "array";
            j._value_valid = true;
            j._jsonString = "[]";
            j._jsonString_valid = false;
            //j._status already = 0
            return j;
        }

        static public iesJSON CreateErrorObject()
        {
            iesJSON j = new iesJSON();
            j._value = null;
            j._jsonType = "error";
            j._value_valid = false;
            j._jsonString = "!error!";
            j._jsonString_valid = false;
            j._status = -1;
            return j;
        }

        static public iesJSON CreateNull()
        {
            iesJSON j = new iesJSON();
            j._value = null;
            j._jsonType = "null";
            j._value_valid = true;
            j._jsonString = "null";
            j._jsonString_valid = true;
            j._status = 0;
            return j;
        }

        // Create a single-node iesJSON with the _value set to singleObj
        //DEFAULT-PARAMETERS
        //static public iesJSON CreateItem(object singleObj) { return CreateItem(singleObj,false); }
        //static public iesJSON CreateItem(object singleObj, bool IncludeStats) {
        static public iesJSON CreateItem(object singleObj, bool IncludeStats = false)
        {
            object o = singleObj;
            iesJSON j = new iesJSON();
            j._jsonType = GetObjType(o);
            if (IncludeStats) { j.TrackStats = true; }
            switch (j._jsonType)
            {
                case "int":
                case "int32":
                case "integer":
                    o = System.Convert.ToDouble((int)o);
                    j._jsonType = "number";
                    break;
                case "int16":
                case "short":
                    o = System.Convert.ToDouble((short)o);
                    j._jsonType = "number";
                    break;
                case "int64":
                case "long":
                    o = System.Convert.ToDouble((long)o);
                    j._jsonType = "number";
                    break;
                case "uint":
                case "uint32":
                    o = System.Convert.ToDouble((uint)o);
                    j._jsonType = "number";
                    break;
                case "ulong":
                case "uint64":
                    o = System.Convert.ToDouble((ulong)o);
                    j._jsonType = "number";
                    break;
                case "float":
                    o = System.Convert.ToDouble((float)o);
                    j._jsonType = "number";
                    break;
                case "decimal":
                    o = System.Convert.ToDouble((decimal)o);
                    j._jsonType = "number";
                    break;
                case "double":
                    o = (double)o;
                    j._jsonType = "number";
                    break;
                case "dbnull":
                    o = null;
                    j._jsonType = "null";
                    break;
            }
            switch (j._jsonType)
            {
                case "string":
                case "number":
                case "boolean":
                case "null":
                    j._value = o;
                    j._value_valid = true;
                    j._status = 0;
                    break;
                default:
                    j.AddStatusMessage("ERROR: Invalid type: '" + j._jsonType + "' [err-47]");
                    j._jsonType = "";
                    j._status = -47;
                    j._value_valid = false;
                    // DEBUG - get info about the DB TYPE
                    // If you encounter an unknown DB type, the debug lines below will
                    // replace the item VALUE with the DB TYPE allow you to see the invalid type.
                    // DO NOT LEAVE THIS CODE IN PRODUCTION
                    //j._value=j._jsonType;
                    //j._jsonType="string";
                    //j._value_valid=true;
                    //j._status=0;
                    break;
            }
            return j;
        }

        // NOTE: The reason we can create a generic "this" enumerator such as...
        //   public object this [int index]
        // is that it would always return a generic object and the C# code would then have to cast that object explicitly.
        // It is better to have the [] get the object and then we can use the .Value function to get/set the value


        // this[index] - where index is an integer
        // WARNING: The "get" function will CREATE AN ITEM in the object/array if it does not exist (with exceptions)
        // Allows numeric index into an iesJSON array or object
        // Base level of the iesJSON must be an "array" or "object"
        public iesJSON this[int index]
        {
            get
            {
                iesJSON o = null;
                try { o = this.ItemAt(index); } catch { }
                if (o == null)
                {
                    try
                    {
                        if (_jsonType == "array") { o = CreateItem(null, TrackStats); AddToArrayBase(o); }
                        // Cannot add an item to an 'object' without a key value
                    }
                    catch { }
                }
                if (o == null) { return CreateErrorObject(); }
                return o;
            }
            set
            {
                // Setting an item using [] requires that you pass in an iesJSON object
                // If needed, you can use the CreateItem() method to generate an iesJSON object to pass to this routine.
                this.ReplaceAt(index, value);
            }
        }

        // this[index] - where index is a string
        // WARNING: The "get" function will CREATE AN ITEM in the object/array if it does not exist
        // Always returns an iesJSON.  If reference is not found, we return an "error" iesJSON object
        public iesJSON this[string indKey]
        {
            get
            {
                iesJSON o = null;
                try { o = this.GetObj(indKey); } catch { }
                if (o == null)
                {
                    iesJSON p = null;
                    object k = null;
                    p = GetParentObj(indKey, ref k);
                    if (p != null)
                    {
                        if (p.jsonType == "object")
                        {
                            string s = k + "";
                            //o=CreateItem("k=" + s);
                            if (s.Trim() != "") { o = CreateItem(null, TrackStats); o.Key = s; p.AddToObjBase(s, o); }
                        }
                        if (p.jsonType == "array")
                        {
                            int i;
                            try
                            {
                                string t = GetObjType(k);
                                switch (t)
                                {
                                    case "integer": i = (int)k; break;
                                    case "double": i = System.Convert.ToInt32((double)k); break;
                                    default: i = int.Parse(k + ""); break;
                                }
                                return p[i];
                            }
                            catch { }
                        }
                    }
                }
                if (o == null) { return CreateErrorObject(); }
                return o;
            }
            set
            {
                // Setting an item using [] requires that you pass in an iesJSON object
                // If needed, you can use the CreateItem() method to generate an iesJSON object to pass to this routine.
                this.ReplaceAt(indKey, value);
            }
        }

        //IEnumerator and IEnumerable require these methods.
        public IEnumerator GetEnumerator()
        {
            return new AEnumerator(this);
        }

        private class AEnumerator : IEnumerator
        {
            public AEnumerator(iesJSON inst)
            {
                this.instance = inst;
            }

            private iesJSON instance;
            public int enum_pos = -1;  // -2=invalid, -1=before first record, 0=first record, etc.

            //IEnumerator
            public bool MoveNext()
            {
                if (instance._status != 0 || !instance._value_valid) { enum_pos = -2; return false; }
                if (enum_pos >= -1) { enum_pos++; }
                if (instance._jsonType != "object" && instance._jsonType != "array")
                {
                    // This is a single type: string, number, null, etc.  There is only one of these.
                    if (enum_pos > 0) { enum_pos = -2; return false; }
                }
                else
                {
                    // jsonType="object" OR "array" - check if we are past the end of the List count.
                    try
                    {
                        if (enum_pos >= ((System.Collections.Generic.List<object>)instance._value).Count) { enum_pos = -2; return false; }
                    }
                    catch
                    {
                        enum_pos = -2;
                        return false;
                    } // In certain unexpected circumstances the above code generated an error
                }
                return true;
            }

            //IEnumerable - NOTE: discovered that Reset() never seems to get called.
            public void Reset()
            { enum_pos = -1; }

            public object Current
            {
                get
                {
                    if (instance._status != 0 || !instance._value_valid) { throw new InvalidOperationException(); }
                    if (enum_pos < 0) { throw new InvalidOperationException(); }
                    if (instance._jsonType != "object" && instance._jsonType != "array")
                    {
                        if (enum_pos > 0) { throw new InvalidOperationException(); }
                        return instance;
                    }
                    // otherwise we are a _jsonType="object" or _jsonType="array"
                    try
                    {
                        return ((System.Collections.Generic.List<object>)instance._value)[enum_pos];  // Have to subtract 1 because the array is 0 based.
                    }
                    catch (IndexOutOfRangeException)
                    {
                        throw new InvalidOperationException();
                    }
                }
            }

            public void Dispose()
            {
                // FUTURE: do we need to do something here?
                instance.Clear();
            }
        } //class AEnum

        // *** Clear()
        // *** src=0 indicates outside source is invalidating us
        // *** src=1 indicates we are invalidating ourself (a routine in this object)
        //DEFAULT-PARAMETERS
        //public void Clear() { Clear(true,0,false); }
        //public void Clear(bool ClearParent) { Clear(ClearParent,0,false); }
        //public void Clear(bool ClearParent,int src) { Clear(ClearParent,src,false); }
        //public void Clear(bool ClearParent,int src,bool bClearStats) {
        //FUTURE: May want to clear any string values in _keep... but need to retain _keep object and _keep parameters
        public void Clear(bool ClearParent = true, int src = 0, bool bClearStats = false)
        {
            if (stats != null)
            {
                if (bClearStats)
                {
                    ClearStats(true);
                }
                else
                {
                    IncStats("stat_Clear");
                    if (src == 0) { IncStats("stat_ClearFromOther"); }
                    if (src == 1) { IncStats("stat_ClearFromSelf"); }
                }
            }
            _status = 0;
            _jsonType = "";
            _value = null;
            _value_valid = false;
            InvalidateJsonString(src); // *** Resets _jsonString and _jsonString_valid (and notifies Parent of changes)
            endpos = 0;
            if (ClearParent) { Parent = null; }
        } // End Clear()

        // *** InvalidateJsonString()
        // *** Indicate that the jsonString no-longer matches the "value" object
        // *** This is done when the "value" of the JSON object changes
        // *** src=0 indicates outside source is invalidating us
        // *** src=1 indicates we are invalidating ourself (a routine in this object)
        // *** src=2 indicates that one of our child objects is invalidating us
        //DEFAULT-PARAMETERS
        //public void InvalidateJsonString() { InvalidateJsonString(0); }
        //public void InvalidateJsonString(int src) {
        public void InvalidateJsonString(int src = 0)
        {
            if (stats != null)
            {
                IncStats("stat_Invalidate");
                if (src == 0) { IncStats("stat_InvalidateFromOther"); }
                if (src == 1) { IncStats("stat_InvalidateFromSelf"); }
                if (src == 2) { IncStats("stat_InvalidateFromChild"); }
            }
            _jsonString = "";
            _jsonString_valid = false;
            // *** if our jsonString is invalid, { the PARENT jsonString would also be invalid
            if (!(Parent == null)) { Parent.InvalidateJsonString(2); }
        } // End InvalidateJsonString()


        // GetStatsBase()
        // If an object is passed in, then all stats are added to that object AND the same object is returned.
        // If a null is passed in, then we create an iesJSON to return.
        // This allows the method to be used as and "Add" or a "Get".
        public string GetStatsBase22() { if (stats == null) { return null; } else { return stats.jsonString; } }

        //DEFAULT PARAMETERS
        //public iesJSON GetStatsBase() { return GetStatsBase(null, false); }
        //public iesJSON GetStatsBase(bool DoNotReturnNull) { return GetStatsBase(null, DoNotReturnNull); }
        public iesJSON GetStatsBase(bool DoNotReturnNull = false) { return GetStatsBase(null, DoNotReturnNull); }

        //DEFAULT PARAMETERS
        //public iesJSON GetStatsBase(iesJSON addToJSON) { return GetStatsBase(addToJSON, false); }
        //public iesJSON GetStatsBase(iesJSON addToJSON, bool DoNotReturnNull) {
        public iesJSON GetStatsBase(iesJSON addToJSON, bool DoNotReturnNull = false)
        {
            iesJSON j;
            if (stats == null)
            {
                if (!DoNotReturnNull) { return null; }
                else { j = CreateEmptyObject(); j.NoStatsOrMsgs = true; return j; }
            }
            j = addToJSON;
            if (j == null) { j = CreateEmptyObject(); j.NoStatsOrMsgs = true; }

            double d; string k;
            foreach (object o in stats)
            {
                iesJSON p;
                p = (iesJSON)o;
                // If this is an numeric statistic, add it together.
                k = (p.Key.ToString()) + "";
                if ((substr(k, 0, 5) == "stat_") && ((p.jsonType) == "number"))
                {
                    d = j.GetDouble(k, 0);
                    j.AddToObjBase(k, d + p.ValueDbl);
                }
                // If this is a message list (error message list) then append the messages
                if (k == "StatusMessages")
                {
                    foreach (object m in ((iesJSON)o))
                    {
                        j.AddToArray("StatusMessages", m);
                    }
                }
            }
            return j;
        } // End GetStatsBase()

        // *** Get Stats for this object AND ALL CHILD iesJSON OBJECTS
        public iesJSON GetStatsAll(ref iesJSON addToJSON)
        {
            iesJSON j;
            j = GetStatsBase(addToJSON);  //If addToJSON is null, this will create a JSON object
            if ((_status != 0) || (!_value_valid)) { return j; } // Cannot iterate through an array/object if it is not valid.
            if ((_jsonType == "object") || (_jsonType == "array"))
            {
                foreach (object o in ((System.Collections.Generic.List<object>)_value))
                {
                    GetStatsAll(ref j); // Adds the stats to j
                }
            }
            return j;
        } // End GetStatsAll()

        // FUTURE:
        // Merge()

        // Make a clone/duplicate of "this" iesJSON (for all sub-items, make a copy of them as well down the entire tree)
        //DEFAULT-PARAMETERS
        //public iesJSON Clone() { return Clone(true, true); }
        //public iesJSON Clone(bool BuildValue) { return Clone(BuildValue, true); }
        //public iesJSON Clone(bool BuildValue,bool BuildString) {
        public iesJSON Clone(bool BuildValue = true, bool BuildString = true)
        {
            if (stats != null) { IncStats("stat_Clone"); }
            iesJSON j = new iesJSON();
            j.UseFlexJson = UseFlexJson;
            j.ALLOW_SINGLE_QUOTE_STRINGS = ALLOW_SINGLE_QUOTE_STRINGS;
            j.SerializeVoidJSON(this, BuildValue, BuildString);
            return j;
        }

        // Clone "this" iesJSON into a specified iesJSON object.
        // Make a clone/duplicate of "this" iesJSON (for all sub-items, make a copy of them as well down the entire tree)
        // NOTE! Clears the destination object first!
        //DEFAULT-PARAMETERS
        //public bool CloneTo(iesJSON toJSONobj) { return CloneTo(toJSONobj,true,true); }
        //public bool CloneTo(iesJSON toJSONobj, bool BuildValue) { return CloneTo(toJSONobj,BuildValue,true); }
        //public bool CloneTo(iesJSON toJSONobj, bool BuildValue,bool BuildString) {
        public bool CloneTo(iesJSON toJSONobj, bool BuildValue = true, bool BuildString = true)
        {
            if (stats != null) { IncStats("stat_CloneTo"); }
            if (toJSONobj == null) { return false; }
            try
            {
                toJSONobj.Clear(); // Make sure there is nothing in the destination object.
                toJSONobj.SerializeVoidJSON(this, BuildValue, BuildString);
            }
            catch { return false; }
            return true;
        }

        // GetDisplayList() - turns a list into an HTML list (separated by <br>)
        //DEFAULT-PARAMETERS
        //public string GetDisplayList() { return GetDisplayList(false); }
        //public string GetDisplayList(bool DisplayJsonAsString) {
        public string GetDisplayList(bool DisplayJsonAsString = false, bool includeSpacingComments = false)
        {
            string ret = "", k, t, z; object g; int c;
            System.Text.StringBuilder s = new System.Text.StringBuilder();
            if (_status != 0) { return ret; }
            if (!ValidateValue()) { return null; } // *** Unable to validate/Deserialize the value
            if ((_jsonType == "object") || (_jsonType == "array"))
            {
                c = 0;
                foreach (object o in this)
                {
                    k = ((iesJSON)o).Key;
                    t = ((iesJSON)o).jsonType;
                    iesJSON kp = ((iesJSON)o)._keep;
                    if (c > 0) { s.Append("<br>\n"); }
                    z = "";
                    switch (t)
                    {
                        case "object":
                        case "array":
                            if (DisplayJsonAsString) { z = ((iesJSON)o).jsonString; }
                            else { z = "*** " + t; }
                            break;
                        default:
                            g = ((iesJSON)o)._value;
                            z = g + "";
                            break;
                    } // end switch
                    if (_jsonType == "object") { s.Append(k + "=" + z); }
                    else { s.Append(z); }
                    // Include Comments and Spacing if specified...
                    if (includeSpacingComments)
                    {
                        if (!(kp == null))
                        {
                            if (kp["keepSpacing"].ToBool() || kp["keepComments"].ToBool())
                            {
                                z = "";
                                z = kp.GetStr("preKey");
                                if (!(z == "")) { s.Append("[preKey]" + EncodeString(z) + "<br>\n"); }
                                z = "";
                                z = kp.GetStr("postKey");
                                if (!(z == "")) { s.Append("[postKey]" + EncodeString(z) + "<br>\n"); }
                                z = "";
                                z = kp.GetStr("preSpace");
                                if (!(z == "")) { s.Append("[preSpace]" + EncodeString(z) + "<br>\n"); }
                                z = "";
                                z = kp.GetStr("postSpace");
                                if (!(z == "")) { s.Append("[postSpace]" + EncodeString(z) + "<br>\n"); }
                                z = "";
                                z = kp.GetStr("finalSpacing");
                                if (!(z == "")) { s.Append("[finalSpacing]" + EncodeString(z) + "<br>\n"); }
                            }
                        }
                    }
                    c = c + 1;
                } // end foreach
                ret = s.ToString();
            }
            else
            { // For all other types other than 'object' or 'array'
                ret = _value + "";
            } // end else/if

            // Include Comments and Spacing if specified...
            if (includeSpacingComments)
            {
                if (!(_keep == null))
                {
                    if (this.keepSpacing || this.keepComments)
                    {
                        ret += "<br>\n";
                        z = "";
                        z = _keep.GetStr("preKey");
                        if (!(z == "")) { s.Append("[preKey]" + EncodeString(z) + "<br>\n"); }
                        z = "";
                        z = _keep.GetStr("postKey");
                        if (!(z == "")) { s.Append("[postKey]" + EncodeString(z) + "<br>\n"); }
                        z = "";
                        z = _keep.GetStr("preSpace");
                        if (!(z == "")) { ret += "[preSpace]" + EncodeString(z) + "<br>\n"; }
                        z = "";
                        z = _keep.GetStr("postSpace");
                        if (!(z == "")) { ret += "[postSpace]" + EncodeString(z) + "<br>\n"; }
                        z = "";
                        z = _keep.GetStr("finalSpacing");
                        if (!(z == "")) { ret += "[finalSpacing]" + EncodeString(z) + "<br>\n"; }
                    }
                }
            }

            return ret;
        } // End GetDisplayList()

        // GetAllParams() - turns an object into an HTML list containing all parameters of the iesJSON item/list (separated by <br>)
        // Used for debuging - data dump of all parameters
        public string GetAllParams(string lineSeparator = "<br>\n", string linePrefix = "")
        {
            System.Text.StringBuilder s = new System.Text.StringBuilder();

            s.Append(linePrefix + "============================ " + _jsonType + " HEADER " + lineSeparator);
            s.Append(linePrefix + "_status:" + _status.ToString() + lineSeparator);
            s.Append(linePrefix + "_statusMsg:\"" + _statusMsg + "\"" + lineSeparator);
            s.Append(linePrefix + "tmpStatusMsg:\"" + tmpStatusMsg + "\"" + lineSeparator);
            s.Append(linePrefix + "_jsonType:" + _jsonType + lineSeparator);
            s.Append(linePrefix + "_key:" + _key + lineSeparator);

            if ((_jsonType == "object") || (_jsonType == "array"))
            {
                s.Append(linePrefix + "_value: <" + _jsonType + "> shown below " + lineSeparator);
            }
            else
            {
                if (_jsonType == "string")
                {
                    s.Append(linePrefix + "_value:\"" + _value + "\"" + lineSeparator);
                }
                else
                {
                    s.Append(linePrefix + "_value:" + _value + lineSeparator);
                }
            }
            s.Append(linePrefix + "_value_valid:" + _value_valid.ToString() + lineSeparator);
            s.Append(linePrefix + "_jsonString:\"" + _jsonString + "\"" + lineSeparator);
            s.Append(linePrefix + "_jsonString_valid:" + _jsonString_valid.ToString() + lineSeparator);
            s.Append(linePrefix + "endpos:" + endpos.ToString() + lineSeparator);
            if (Parent == null) { s.Append(linePrefix + "Parent: <null>" + lineSeparator); }
            else { s.Append(linePrefix + "Parent: <object>" + lineSeparator); }
            s.Append(linePrefix + "ALLOW_SINGLE_QUOTE_STRINGS:" + ALLOW_SINGLE_QUOTE_STRINGS.ToString() + lineSeparator);
            s.Append(linePrefix + "ENCODE_SINGLE_QUOTES:" + ENCODE_SINGLE_QUOTES.ToString() + lineSeparator);
            s.Append(linePrefix + "_NoStatsOrMsgs:" + _NoStatsOrMsgs.ToString() + lineSeparator);
            try
            {
                s.Append(linePrefix + "_UseFlexJson:" + _UseFlexJson.ToString() + lineSeparator);
                if (_keep == null)
                {
                    s.Append(linePrefix + "_keep: <null> " + lineSeparator);
                }
                else
                {
                    s.Append(linePrefix + "------------------------- _keep:" + lineSeparator);
                    s.Append(_keep.GetAllParams(lineSeparator, linePrefix + ">>"));
                }

                if (stats == null)
                {
                    s.Append(linePrefix + "stats: <null> " + lineSeparator);
                }
                else
                {
                    s.Append(linePrefix + "------------------------- stats:" + lineSeparator);
                    s.Append(stats.GetAllParams(lineSeparator, linePrefix + ">>"));
                }

                if ((_jsonType == "object") || (_jsonType == "array"))
                {
                    s.Append(linePrefix + "------------------------- " + _jsonType + ":" + lineSeparator);

                    foreach (object o in this)
                    {
                        s.Append(((iesJSON)o).GetAllParams(lineSeparator, linePrefix + ">>"));
                    } // end foreach
                }
            }
            catch (Exception e9) { s.Append(linePrefix + "*ERROR*" + lineSeparator); }

            return s.ToString();
        } // End GetAllParams()

        public bool UseFlexJson
        {
            get { return _UseFlexJson; }
            set
            {
                _UseFlexJson = value;
                InvalidateJsonString(1);  // If we switch to FlexJson then all our stored JSON strings could be incorrect.
            }
        }

        public string jsonString
        {
            get
            {
                if (stats != null) { IncStats("stat_jsonString_get"); }
                if (_status != 0) { return (""); } // *** ERROR - status is invalid
                if (!_jsonString_valid)
                {
                    if (!_value_valid)
                    {
                        // *** Neither is valid - this JSON object is null
                        return "";
                    }
                    else
                    {
                        // *** First we need to serialize
                        SerializeMe();
                        if (_status != 0) { return ""; } // *** ERROR - status is invalid
                        return _jsonString;
                    }
                }
                else
                {  // jsonString was already valid.
                    return _jsonString;
                }
            }
            set
            {
                if (stats != null) { IncStats("stat_jsonString_set"); }
                if ((_status == 0) && (_jsonString_valid) && (_jsonString == value))
                {
                    return;
                }
                else
                {
                    this.Clear(false, 1);  // *** This clears _value AND notifies Parent that we have changed.
                    _jsonString = value;
                    _jsonString_valid = true;  // *** This does not indicate 'valid JSON syntax' but only that the string has been populated.
                }
            } //End Set
        } // End Property jsonString()


        public bool keepSpacing
        {
            get
            {
                if (_keep == null) { return false; }
                return _keep.GetBool("keepSpacing");
            }
            set
            {
                if (value) { keepSpacingAndComments(1, -1); }
                else { keepSpacingAndComments(0, -1); }
            }
        }

        public bool keepComments
        {
            get
            {
                if (_keep == null) { return false; }
                return _keep.GetBool("keepComments");
            }
            set
            {
                if (value) { keepSpacingAndComments(-1, 1); }
                else { keepSpacingAndComments(-1, 0); }
            }
        }

        // keepSpacingAndComments() - setsup the spacing/comments object used for Flex config files
        //   flag values: -1 leave default value, 0 Set to FALSE, 1 Set to TRUE
        public void keepSpacingAndComments(int spacing_flag = -1, int comments_flag = -1)
        {
            if (_keep == null)
            {
                // Need to create object - NOTE: Default settings are FALSE for both keepSpacing and keepComments
                _keep = new iesJSON("{\"keepSpacing\":false,\"keepComments\":false}");
            }
            if (spacing_flag == 0) { _keep["keepSpacing"].Value = false; }
            if (spacing_flag == 1) { _keep["keepSpacing"].Value = true; }
            if (comments_flag == 0) { _keep["keepComments"].Value = false; }
            if (comments_flag == 1) { _keep["keepComments"].Value = true; }
        }

        public void keepSpacingAndComments(bool spacing_flag, bool comments_flag)
        {
            keepSpacingAndComments(Convert.ToInt32(spacing_flag), Convert.ToInt32(comments_flag));
        }

        // addSpacingOrComment()
        // NOTE! Spacing MUST BE white space: space, tab, or carriage return
        // NOTE! Comments must already be properly marked for example /* comment */
        // NOTE: This routine uses deserialize to verify that what is being added is nothing but white space and comments?
        //        Can override this feature by setting verifyFormat=true, but this can lead to unwanted behavior!
        // Return: true=Spacing/Comment has been added.
        public bool addSpacingOrComment(string addString, bool addAfter = true, bool verifyFormat = true, bool addToKey = false)
        {
            // If verifyFormat=true, then we deserialize the Spacing/Comment to make sure it meets formatting requirement
            if (verifyFormat)
            {
                try
                {
                    iesJSON verifyObj = new iesJSON();
                    verifyObj.UseFlexJson = true;
                    verifyObj.keepSpacingAndComments(1, 1);
                    verifyObj.Deserialize(addString);
                    if (verifyObj.Status != 0)
                    {
                        return false; // bad formatting
                    }
                    // Parsed JSON must be either an empty string or NULL
                    if (verifyObj._jsonType == "string")
                    {
                        if (verifyObj.ToStr() != "")
                        {
                            return false; // field contained invalid formatting
                        }
                    }
                    else
                    {
                        if (verifyObj._jsonType != "null")
                        {
                            return false; // field contained invalid formatting
                        }
                    }
                }
                catch (Exception verifyErr)
                {
                    return false; // error occurred - invalid formatting
                }
            }

            // Check to see if _keep is already defined - if not, create it
            if (_keep == null)
            {
                this.keepSpacingAndComments(); // leave defaults
            }

            // Get current pre/post spacing and comments
            string currKeep = null;
            string addToToken = null;
            if (addAfter)
            {
                addToToken = "post";  // Add Comment/Spacing AFTER current object (postKey or postSpace)
            }
            else
            {
                addToToken = "pre";  // Add Comment/Spacing BEFORE current object (preKey or preSpace)
            }
            if (addToKey == false)
            {
                addToToken += "Space";
            }
            else
            {
                addToToken += "Key";
            }

            // Add Comment/Spacing to current object
            currKeep = _keep[addToToken].ToStr();
            currKeep += addString;
            _keep[addToToken].Value = currKeep;

            return true;
        }


        public int LineNumber  // Indicates on which line this item was found during parse
        {
            get
            {
                if (_keep == null) { return -1; }
                return _keep["LineNumber"].ToInt(-1);
            }
        }

        public int LinePosition // Indicates at which column position on the line this item was found
        {
            get
            {
                if (_keep == null) { return -1; }
                return _keep["LinePosition"].ToInt(-1);
            }
        }

        public int AbsolutePosition()  // Indicates at which character position in the file this item was found (absolute position regardless of carriage returns)
        {
        get:
            if (_keep == null) { return -1; }
            return _keep["AbsolutePosition"].ToInt(-1);
        }


        // ItemAt returns an iesJSON node from a specific index which can be an: object, array, string, number, bool, or null
        // This is the same as iesJSON[index] where index is an integer
        public iesJSON ItemAt(int index)
        {
            if (!ValidateValue()) { return null; } // *** Unable to validate/Deserialize the value
            else
            {
                if (_jsonType == "object" || _jsonType == "array")
                {
                    // This returns a iesJSON object from the List (array)
                    // If the index is out of range, an error will be raised.
                    try
                    {
                        return (iesJSON)((System.Collections.Generic.List<object>)_value)[index];
                    }
                    catch { return null; }
                }
                // If this is not an object/array then...
                if (index == 0) { return this; } // For string, number, bool, null index=0 we can just return this object
                return null;
            }
        }

        //DEFAULT-PARAMETERS
        //public bool ReplaceAt(int index, iesJSON newItem) { return ReplaceAt(index,newItem,false); }
        //public bool ReplaceAt(int index, iesJSON newItem, bool OverrideKey) {
        public bool ReplaceAt(int index, iesJSON newItem, bool OverrideKey = false)
        {
            if (!ValidateValue()) { AddStatusMessage("ERROR: Unable to validate/Deserialize the value. [err443]"); return false; } // *** Unable to validate/Deserialize the value
            else
            {
                InvalidateJsonString(1); // Whether we succeed or not, flag the new value and all parents as having changed.
                if (_jsonType == "object" || _jsonType == "array")
                {
                    try
                    {
                        newItem.Parent = this;
                        ((System.Collections.Generic.List<object>)_value)[index] = newItem;
                    }
                    catch { return false; }
                }
                else
                {
                    if (index == 0)
                    {
                        // Replace this item VALUE
                        _value = newItem.Value;
                        _jsonType = newItem.jsonType;
                        if (OverrideKey) { _key = newItem.Key; }  // USUALLY WE DO NOT CHANGE THE KEY - THIS IS ID WITHIN THE PARENT OBJECT.
                        _value_valid = true;
                    }
                    else
                    {
                        // Index is invalid - must be 0 for a non-array/non-object
                        return false;
                    }
                }
            }
            return true;
        }

        // FUTURE: we can end up in an infinite loop if index does contain a "." but it is inside quotes which makes it a JSON reference.
        public bool ReplaceAt(string index, iesJSON newItem)
        {
            if (!ValidateValue()) { AddStatusMessage("ERROR: Unable to validate/Deserialize the value. [err445]"); return false; } // *** Unable to validate/Deserialize the value
            else
            {
                if (String.IsNullOrEmpty(index)) { return ReplaceAt(0, newItem); }
                if (index.IndexOf(".") < 0)
                {
                    // Does not contain a ".", so this must be a referencing a "Key"
                    if (_jsonType != "object") { return false; }
                    int i = this.IndexOfKey(index);
                    newItem.Key = index;  // Set the key value of the new item.
                    if (i < 0)
                    {
                        this.AddToObjBase(index, newItem);
                    }
                    else { this.ReplaceAt(i, newItem); }
                }
                else
                {
                    // First determine parent object
                    iesJSON p; object k = null;
                    p = this.GetParentObj(index, ref k);
                    if (p == null) { return false; }
                    p.ReplaceAt(k + "", newItem); // FUTURE: what do we do if k is numeric?
                }
            }
            return true;
        }

        public string Key
        {
            get
            {
                if (_status != 0) { return null; }
                if (!ValidateValue()) { return null; }  // *** Unable to validate/Deserialize the value
                return _key;
            }
            set
            {
                _key = value;
            }
        }

        // Value returns the VALUE from a Key/Value pair.  (if this.jsonType is object or array, then it return a null.  For those two, use the Item property above.)
        public object Value
        {
            get
            {
                if (stats != null) { IncStats("stat_Value_get"); }
                if (!ValidateValue()) { return null; }  // *** Unable to validate/Deserialize the value
                                                        // If object or array, return null
                if (_jsonType == "object" && _jsonType == "array") { return null; }
                return _value;
            } //End get
            set
            {
                if (stats != null) { IncStats("stat_Value_set"); }
                iesJSON o = CreateItem(value);
                this.ReplaceAt(0, o);
            }
        } // End Property Value()

        // Returns the _VALUE from this instance as a String.  Number/Boolean will be converted to a string.  Null/Object/Array will all return "".

        public string ValueString
        {
            get { return this.ToStr(""); }
            set { Value = value; } //Value->set creates the item, sets the jsonType, and invalidates the jsonString
        }

        // CString() is provided here for backwards compatibility
        //DEFAULT-PARAMETERS
        //public string CString() { return ToStr("");}
        //public string CString(string sDefault) { return ToStr(sDefault); }
        public string CString(string sDefault = "") { return ToStr(sDefault); }

        //DEFAULT-PARAMETERS
        //public string ToStr() { return ToStr("");}
        //public string ToStr(string sDefault) {
        public string ToStr(string sDefault = "")
        {  // NOTE! Changed this from "ToString" to "ToStr" (Meaning Convert-to-String) because several times the compiler mixed up this function (named "ToString") with string.ToString() or object.ToString() methods
            if (stats != null) { IncStats("stat_ValueString_get"); }
            if (!ValidateValue()) { return sDefault; } // Error - invalid value - return default
            if (_status != 0 || _jsonType == "object" || _jsonType == "array") { return sDefault; }
            if (_jsonType == "boolean") { return _value.ToString().ToLower(); }
            if (_jsonType == "string") { return (string)_value; }
            if (_jsonType == "number") { return ((double)_value).ToString(); }
            return sDefault;
        } // End Property ValueString()

        // Returns the _VALUE from this instance as an Integer. Number/String will be converted to an Integer.  Null/Object/Array will all return 0.  True=1, False=0
        public int ValueInt
        {
            get { return ToInt(); }
            set { Value = value; }
        }
        //DEFAULT-PARAMETERS
        //public int ToInt() { return ToInt(0); }
        //public int ToInt(int nDefault) {
        public int ToInt(int nDefault = 0)
        {
            if (stats != null) { IncStats("stat_ValueInt_get"); }
            if (!ValidateValue()) { return nDefault; } // Error - invalid value - return false
            if (_status != 0 || _jsonType == "object" || _jsonType == "array") { return nDefault; }
            if (_jsonType == "boolean")
            {
                if ((bool)_value) { return (int)1; }
                else { return (int)0; }
            }
            if (_jsonType == "string")
            {
                try { return System.Convert.ToInt32((string)_value); } catch { }
                return nDefault;
            }
            if (_jsonType == "number") { return System.Convert.ToInt32((double)_value); }
            return nDefault;
        } // End Property ValueInt()

        // Returns the _VALUE from this instance as a double.  Number/String will be converted to a Double.  Null/Object/Array will all return 0.  True=1, False=0
        public double ValueDbl
        {
            get { return ToDbl(); }
            set { Value = value; }
        }
        //DEFAULT-PARAMETERS
        //public double ToDbl() { return ToDbl(0d); }
        //public double ToDbl(double dDefault) {
        public double ToDbl(double dDefault = 0d)
        {
            if (stats != null) { IncStats("stat_ValueDbl_get"); }
            if (!ValidateValue()) { return dDefault; } // Error - invalid value - return false
            if (_status != 0 || _jsonType == "object" || _jsonType == "array") { return dDefault; }
            if (_jsonType == "boolean") { return System.Convert.ToDouble((bool)_value); }
            if (_jsonType == "string")
            {
                try { return System.Convert.ToDouble((string)_value); } catch { }
                return dDefault;
            }
            if (_jsonType == "number") { return (double)_value; }
            return dDefault;
        } // End Property ValueDbl()

        // Returns the _VALUE from this instance as a boolean.
        public bool ValueBool
        {
            get { return ToBool(); }
            set { Value = value; }
        }
        //DEFAULT-PARAMETERS
        //public bool ToBool() { return ToBool(false); }
        //public bool ToBool(bool bDefault) {
        public bool ToBool(bool bDefault = false)
        {
            if (stats != null) { IncStats("stat_ValueBool_get"); }
            if (!ValidateValue()) { return bDefault; } // Error - invalid value - return false
            if (_status != 0 || _jsonType == "object" || _jsonType == "array") { return bDefault; }
            if (_jsonType == "boolean") { return (bool)_value; }
            if (_jsonType == "string")
            {
                string sval = ((string)_value).ToLower();
                if (sval == "true" || sval == "t" || sval == "1" || sval == "-1") { return true; } else { return false; }
            }
            if (_jsonType == "number")
            {
                if ((double)_value != 0d) { return true; } else { return false; }
            }
            return bDefault;
        }

        // ToArray() - Get the current value in the form of an Array - non-Destructive
        // If item is already an array it returns "this" (full object)
        // If item is a null, a null array is returned (array with no elements)
        // If item is an Object and...
        //     flagConvertObjects=false then the object is returned as an object
        //     flagConvertObjects=true then the object is converted to an array
        //         WARNING: MAY CLONE THE ITEM TO KEEP FROM MODIFYING PARENT
        // If item is an element (string,int,etc) it returns an array containing this single element
        // NOTE! This function should not be confused with ConvertToArray (which changes the iesJSON object. ie Destructive)
        //
        // cloneArray causes an Array or Object to be cloned into an Array (so that if it gets modified down the pipe line, you are not destroying the parent)
        public iesJSON ToArray(bool flagConvertObjects = false, bool cloneArray=false)
        {
            switch (this.jsonType)
            {
                case "array":
                    if (cloneArray) {
                        return this.Clone();
                    }
                    return this;
                case "object":
                    if (flagConvertObjects || cloneArray)
                    {
                        iesJSON newArray=this.Clone();
                        newArray.ConvertToArray();
                        return newArray;
                    }
                    return this;
                case "null":
                    iesJSON returnArray2 = new iesJSON("[]");
                    return returnArray2;
                default:
                    iesJSON returnArray3 = new iesJSON("[]");
                    returnArray3.AddToArrayBase(this);
                    return returnArray3;
            }
        }

        // ToArrayWithSplit() - Return an array or a SPLIT of a string
        // Used for FLEX JSON config files so that the file can contain
        //   either an Array or a Comma separated list in string format
        // eg. the following two lines produce the same result...
        //   fields:[id,title,status]
        //   fields:"id,title,status"
        //
        // Also, note that the following line still returns an array with a single item
        //   fields:"id"
        public iesJSON ToArrayWithSplit(char SplitChar=',', bool cloneArray=false) {
            switch (this.jsonType)
            {
                case "string":
                    iesJSON newArray = new iesJSON("[]");
                    foreach(string item in this.ToStr().Split(SplitChar)) {
                        newArray.Add(item);
                    }
                    return newArray;
                default:
                    return this.ToArray(true,cloneArray);
            }
        }

        // Value returns the List<object> from a Key/Value pair.  (if this.jsonType is object or array, then it returns the List.  All other types returns null.)
        public object ValueArray
        {
            get
            {
                if (stats != null) { IncStats("stat_ValueArray_get"); }
                if (!ValidateValue())
                {
                    return null;  // *** Unable to validate/Deserialize the value
                }
                else
                {
                    // If object or array, return null
                    if (_jsonType == "object" && _jsonType == "array") { return _value; }
                    return null;
                }
            } //End get
              // *** NO "SET"
        } // End Property Value()

        // This routine will Deserialize the JsonString if needed in order to make sure the _value is valid.
        public bool ValidateValue()
        {
            if (_status != 0) { return false; } // *** ERROR - status is invalid
            if (!_value_valid)
            {
                if (!_jsonString_valid)
                {
                    // *** Neither is valid - this JSON object is null - Cannot validate value
                    return false;
                }
                else
                {
                    // *** First we need to Deserialize
                    DeserializeMe();
                    if (_status != 0) { return false; } // *** ERROR - status is invalid - failed to Deserialize
                    return true;
                }
            }
            else
            {
                return true;
            }
        }

        public bool Contains(string Search, bool DotNotation = true)
        {
            string[] searchList;
            int k = -99;  // tells loop not to perform 'next' function
            iesJSON next = this;
            if (DotNotation)
            {
                searchList = Search.Split('.');
            }
            else
            {
                searchList = new string[] { Search };
            }
            foreach (string sItem in searchList)
            {
                if (k >= 0)
                {
                    next = next[k];
                }
                string s = sItem.Trim();
                k = -1;
                if (next._jsonType == "object") { k = next.IndexOfKey(s); }
                else if (next._jsonType == "array") { k = next.IndexOf(s); }
                //else { k = -1; }

                if (k < 0) { break; }
            }
            if (k < 0) { return false; }
            return true;
        }

        public int RenameItem(string OldName, string NewName) //return -1 for error or number of fields replaced
        {
            if (OldName == null || NewName == null) { return -1; }
            string old = OldName.ToLower();
            string newn = NewName.ToLower();

            if (!ValidateValue()) { return -1; }
            if (_jsonType != "object") { return -1; }  // Future: do we want to search for values here?  probably not.  Implement IndexOf for search of values?
            int k = 0;
            // NOTE: Here we do a linear search.  FUTURE: if list is sorted, do a binary search.
            foreach (object o in ((System.Collections.Generic.List<object>)_value))
            {
                iesJSON i = (iesJSON)o;
                if (i._key.ToLower() == newn)
                {
                    return -2; //this indicates that newname already exists
                }
            }

            foreach (object o in ((System.Collections.Generic.List<object>)_value))
            {
                iesJSON i = (iesJSON)o;
                try
                {
                    if (i._key.ToLower() == old)
                    {
                        k++;
                        i._key = NewName;
                        this.InvalidateJsonString(1);
                    }
                }
                catch { }
            }
            return k;
        }

        public int IndexOfKey(string Search)
        {
            if (Search == null) { return -1; }
            string s = Search.ToLower();
            if (!ValidateValue()) { return -1; }
            if (_jsonType != "object") { return -1; }  // Future: do we want to search for values here?  probably not.  Implement IndexOf for search of values?
            int k = 0;
            // NOTE: Here we do a linear search.  FUTURE: if list is sorted, do a binary search.
            foreach (object o in ((System.Collections.Generic.List<object>)_value))
            {
                try { if (((iesJSON)o)._key.ToLower() == s) { return k; } } catch { }
                k++;
            }
            return -1;
        }

        public int IndexOf(string Search)
        {
            if (Search == null) { return -1; }
            string s = Search.ToLower();
            if (!ValidateValue()) { return -1; }
            if (_jsonType != "array") { return -1; }  // For objects use IndexOfKey()
            int k = 0;
            // NOTE: Here we do a linear search.  FUTURE: if list is sorted, do a binary search.
            foreach (object o in ((System.Collections.Generic.List<object>)_value))
            {
                try { if (((iesJSON)o)._value.ToString().ToLower() == s) { return k; } } catch { }
                k++;
            }
            return -1;
        }

        // Turns on/off the tracking of stats
        public bool TrackStats
        {
            get
            {
                if (stats != null)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }

            set
            {
                if ((value) && (!_NoStatsOrMsgs))
                {
                    if (stats == null)
                    {
                        // turning on stats
                        stats = CreateEmptyObject();
                        stats.NoStatsOrMsgs = true;
                    }
                }
                else
                {
                    // turning off stats
                    stats = null;
                }
            }
        } // End Property TracksStats()

        // Set this flag to TRUE to keep the object from tracking stats or error messages.
        // It is very important to set this to TRUE for the stats object (to avoid recursive loops).
        public bool NoStatsOrMsgs
        {
            get
            {
                return _NoStatsOrMsgs;
            }
            set
            {
                _NoStatsOrMsgs = value;
                if (value)
                {
                    this.TrackStats = false;
                }
            }
        }

        // Gets Array.Length or Dictionary.Count for array/object.  Gets 1 for all other types (string, number, null, etc.)  Returns 0 if _value is not valid.
        public int Length
        {
            get
            {
                if (_status != 0 || !_value_valid) { return 0; }
                if (_jsonType == "null") { return 0; }
                if (_jsonType != "array" && _jsonType != "object") { return 1; }
                // otherwise we are a _jsonType="object" or _jsonType="array"
                try
                {
                    return ((System.Collections.Generic.List<object>)_value).Count;
                }
                catch { return -1; } //Error occurred... return -1 to indicate an error
            }
            // There is no "set" for length
        }

        public bool jsonString_valid
        {
            get
            {
                return _jsonString_valid;
            }
            // We should not be setting this value.  To create a valid Json string, call SerializeMe()
        } // End Property jsonString_valid()

        public int Status
        {
            get
            {
                return (_status);
            }
        } // End Property

        // Typically this should not be used - however there are a few cases where the owner of the object would like to invalidate the object by setting the status to invalid.
        public void InvalidateStatus(int newStatus = -999)
        {
            _status = newStatus;
            if (newStatus == 0) { _status = -999; }
        }

        public string StatusMessage
        {  // Gets the most recent error message.  stats contains a list of error messages.
            get { return _statusMsg; }
        } // End Property

        public string jsonType
        {
            get
            {
                return _jsonType;
            }
        } // End Property

        public int jsonTypeEnum
        {
            get
            {
                //Return a number representing the JSONtype
                //NOTE: The NUMBER VALUES are critical to the SORT() routine
                //  So that iesJSON arrays get sorted in this order: NULL, Number, String, Boolean, Array, Object, Error
                switch (_jsonType)
                {
                    case "null":
                        return jsonTypeEnum_null;
                    case "number":
                        return jsonTypeEnum_number;
                    case "string":
                        return jsonTypeEnum_string;
                    case "boolean":
                        return jsonTypeEnum_boolean;
                    case "array":
                        return jsonTypeEnum_array;
                    case "object":
                        return jsonTypeEnum_object;
                    case "error":
                        return jsonTypeEnum_error;
                    default: // This should never occur - invalid jsonType
                        return jsonTypeEnum_invalid;
                } // end switch
            } // end get
        }

        /* ConvertToArray()
		** Convert an 'object' to an 'array' - Here we cheat - because a JSON ARRAY is not any different than a JSON OBJECT - only we ignore the field names
		** NOTE: Can only be performed on an OBJECT (or ARRAY which does nothing)
		** RETURN: 0=already is an array, 1=Success, -1=ERROR: item was not an array or object
		**/
        public int ConvertToArray()
        {
            if (this._jsonType == "object")
            {
                _jsonType = "array";
                return 1;
            }
            if (this._jsonType == "array")
            {
                return 0;
            }
            return -1;
        }

        // *** Deserialize()
        // ***   if this item starts with a { then it is an parameter list and must end with a } and anything past that is ignored
        // ***   if this item starts with a [ then it is an array and must end with a ] and anything past that is ignored
        // ***   Other options are String (must be surrounded with quotes), Integer, Boolean, or Null
        // ***   null (all white space) creates an error
        // ***
        // *** returns 0=Successful, anything else represents an invalid status
        // ***
        //DEFAULT-PARAMETERS
        //public int Deserialize(string snewString) { return Deserialize(snewString, 0, false); }
        //public int Deserialize(string snewString, int start) { return Deserialize(snewString, start, false); }
        //public int Deserialize(string snewString, int start, bool OkToClip) {
        public int Deserialize(string snewString, int start = 0, bool OkToClip = false)
        {
            if (stats != null) { IncStats("stat_Deserialize"); }
            //tmpStatusMsg=tmpStatusMsg + "DEBUG: Deserialize<br>";
            Clear(false, 1);
            _jsonString = snewString;
            _jsonString_valid = true;

            return DeserializeMe(start, OkToClip);
        } // End Function

        //DEFAULT-PARAMETERS
        //public int DeserializeMe() { return DeserializeMe(0, false); }
        //public int DeserializeMe(int start) { return DeserializeMe(start, false); }
        //public int DeserializeMe(int start, bool OkToClip) {
        public int DeserializeMe(int start = 0, bool OkToClip = false)
        {
            string c;
            bool keepSP = false;
            bool keepCM = false;
            StringBuilder getSpace = null;
            if (stats != null) { IncStats("stat_DeserializeMe"); }

            if (!_jsonString_valid)
            {
                // *** We can only Deserialize if the jsonString is filled in.  return error.
                _status = -99;
                AddStatusMessage("ERROR: Cannot DeserializeMe() because jsonString is not set.  [err-99]");
                return _status;
            }

            _status = 0;
            _key = null; // *** Default
            _value = null; // *** Default
            _value_valid = false; // *** In case of error, default is invalid
            endpos = start;
            if (_keep != null)
            {
                keepSP = this.keepSpacing;
                keepCM = this.keepComments;
                getSpace = new StringBuilder();
            }
            findnext(keepSP, keepCM, ref getSpace);
            //tmpStatusMsg+="[DeserializeMe.findnext][keepSP=" + keepSP.ToString() + "][keepCM=" + keepCM.ToString() + "]"; // DEBUG
            //if (_keep!=null) { tmpStatusMsg+="[preSpace=" + _keep.GetStr("preSpace") + "]"; } // DEBUG
            //if (getSpace!=null) {tmpStatusMsg+="[getSpace=" + getSpace.ToString() + "]";} // DEBUG
            if (keepSP || keepCM)
            {
                AddSpaceAndClear(_keep, "preSpace", ref getSpace);
                AddLinePosition(_keep, endpos);
            }

            // *** Check for zero length string (null)
            if (endpos >= _jsonString.Length)
            {
                _jsonType = "string"; // FUTURE: Why is this not a null?
                _value = "";
                _value_valid = true;
                goto Done_Deserialization;
            }

            // *** Check for Array or Object or String
            c = substr(_jsonString, endpos, 1);
            if (c == "{") { DeserializeObject(keepSP, keepCM); }
            if (c == "[") { DeserializeArray(keepSP, keepCM); }
            if (c == "\"") { DeserializeMeAsString(c); }
            if (ALLOW_SINGLE_QUOTE_STRINGS && c == "'") { DeserializeMeAsString(c); }
            if (_jsonType != "") { goto Done_Deserialization; }

            // *** Check to see if it is a Boolean value: true or false
            if (substr(_jsonString, endpos, 4).ToLower() == "true")
            {
                _jsonType = "boolean";
                _value = true;
                _value_valid = true;
                endpos = endpos + 4;
                goto Done_Deserialization;
            }
            if (substr(_jsonString, endpos, 5).ToLower() == "false")
            {
                _jsonType = "boolean";
                _value = false;
                _value_valid = true;
                endpos = endpos + 5;
                goto Done_Deserialization;
            }

            // *** Check to see if it is a NULL value
            if (substr(_jsonString, endpos, 4).ToLower() == "null")
            {
                _jsonType = "null";
                _value = null;
                _value_valid = true;
                endpos = endpos + 4;
                goto Done_Deserialization;
            }

            // Check to see if this is a FLEX JSON string or a NUMBER
            findstring();
            //findnumber();

            c = substr(_jsonString, start, endpos - start).Trim();
            if (c == "")
            {
                _status = -11;
                //THIS IS NOT NECESSARILY AN ERROR! MAY BE A NORMAL END TO A SECTION OF THE JSON STRING
                _value = null;
                _jsonType = "null";
                goto Done_Deserialization;
            }

            // *** determine if this is a number by trying to parse it...
            try
            {
                double valueCheck;
                if (Double.TryParse(c, out valueCheck))
                {
                    _value = valueCheck;

                    // It worked... keep going...
                    _jsonType = "number";
                    _value_valid = true;
                    goto Done_Deserialization;
                }
                else
                {
                    // value isn't a double, consider it a string
                    _value = c;
                    _jsonType = "string";
                    _value_valid = true;
                    goto Done_Deserialization;
                }
            }
            catch
            {
                if (_UseFlexJson)
                {
                    _value = c;
                    _jsonType = "string";
                    _value_valid = true;
                    goto Done_Deserialization;
                }
                _value = null; //parse did not work
            }
            // *** Was not a number or string - error
            _status = -12; // *** Invalid string/format found
            AddStatusMessage("Invalid Format - String/Number/Boolean/Null not found [c=" + c + ",flex=" + _UseFlexJson.ToString() + "]");

        //	Done_StringNumber:
        //		if (substr(_jsonString,endpos,2).ToLower()=="/*") {
        //			findcommentend();
        //			// DO STUFF HERE - FUTURE
        //			}

        Done_Deserialization:
            // *** Check if there is anything left at the end of the string... if so, then this was not valid JSON (or we are mid-string in parsing the json string)
            if (_status == 0 || _status == -11)
            {
                _value_valid = true;
                if (OkToClip)
                {
                    _jsonString = substr(_jsonString, start, endpos - start);
                }
                else
                {
                    findnext(keepSP, keepCM, ref getSpace);  // This not only removes remaining spaces, but also skips comments if it is Flex Json
                    if (keepSP || keepCM)
                    {
                        AddSpaceAndClear(_keep, "postSpace", ref getSpace);
                    }
                    if (!(endpos > _jsonString.Length))
                    {
                        if (substr(_jsonString, endpos).Trim() != "") { _status = -1; } // *** Indicates that the json syntax is wrong - there are characters past the end of the JSON - if you don't care, then set the OkToClip flag
                                                                                        // If keep spacing/comment flag set, then store the remainder
                        if (keepSP || keepCM)
                        {
                            //_keep["finalSpace"].Value=substr(_jsonString,endpos);
                            AddSpace(_keep, "finalSpace", substr(_jsonString, endpos));
                        }
                    }
                }
            }
            else
            {
                _jsonString = ""; // safety to make sure we do not recusively try to re-interpret the json string
            }
            // *** WE MUST LEAVE endpos FOR REFERENCE!  (this is the endpos of the ORIGINAL string... not the clipped end-result
            return (_status);
        } // End Function

        private void DeserializeMeAsString(string endChar)
        {
            int dsErr = 0; string dsErrMsg = "";
            _value = DeserializeString(_jsonString, endChar, 1, ref endpos, ref dsErr, dsErrMsg);
            if (dsErr == 0)
            {
                _jsonType = "string";
                _value_valid = true;
            }
            else { StatusErr(dsErr, dsErrMsg); }
        }

        static public string DeserializeString(string js, string endChar, int FirstCharFlag)
        {
            int iPos = 0, dsErr = 0; string dsErrMsg = "";
            return DeserializeString(js, endChar, FirstCharFlag, ref iPos, ref dsErr, dsErrMsg);
        }

        static public string DeserializeString(string js, string endChar, int FirstCharFlag, ref int pos)
        {
            int dsErr = 0; string dsErrMsg = "";
            return DeserializeString(js, endChar, FirstCharFlag, ref pos, ref dsErr, dsErrMsg);
        }

        // DeserializeString()
        // FirstCharFlag: 0=There is no quote character to skip, 1=Skip first character which is a quote, 2=Use first character as the endChar
        // endChar: ""=do not look for an end character (process entire string).  If endChar is set to a character, look for that character as an end of string. (usually a quote)
        static public string DeserializeString(string js, string endChar, int FirstCharFlag, ref int pos, ref int dsErr, string dsErrMsg)
        {
            int st; string c; bool brk = false; string endChar2;
            System.Text.StringBuilder w = new System.Text.StringBuilder();

            // *** For now we just do a quick search for end of string.  Future we need to expand for special escape characters.
            if (FirstCharFlag == 2) { endChar2 = substr(js, pos, 1); }  // *** If specified, get first quote symbol
            else { endChar2 = endChar; }
            if (FirstCharFlag >= 1) { pos = pos + 1; }  // *** IMPORTANT! Move past first quote symbol
            st = pos;
            do
            {
                if (pos > js.Length)
                {
                    if (endChar2 != "")
                    {
                        dsErr = -33;
                        dsErrMsg = "Found end of the buffer before finding the end quote @Line:" + currLine(pos, ref js) + " @Position:" + currLinePos(pos, ref js) + " (e-33)";
                        // We still return the string even though it is missing the final quote.
                    }
                    break;
                }
                c = substr(js, pos, 1);
                if (c == endChar2)
                {
                    // *** End quote found
                    pos = pos + 1;
                    break;
                }
                else
                {
                    switch (c)
                    {
                        case "\\":
                            // *** Turn this into an escape character!
                            pos = pos + 1;  // *** get next character
                            c = substr(js, pos, 1);
                            switch (c)
                            {
                                case "\\":
                                case "/":
                                case "'":
                                case "\"":
                                    w.Append(c);
                                    break;
                                case "t":
                                    w.Append("\t"); //Tab character
                                    break;
                                case "b":
                                    w.Append(Convert.ToChar(8));
                                    break;
                                case "c":
                                    w.Append(Convert.ToChar(13));
                                    break;
                                case "n":
                                    w.Append("\n");  //New Line Character
                                    break;
                                case "r":
                                    w.Append("\r");  //LineFeedCarriageReturn
                                    break;
                                case "v":
                                    w.Append("*");  // ***** FUTURE: Need to determine this character!
                                    break;
                                case "u":
                                    // *** Here we need to get the next 4 digits and turn them into a character
                                    c = substr(js, pos + 1, 4);
                                    if (IsHex4(c))
                                    {
                                        string asciiValue = System.Convert.ToChar(System.Convert.ToUInt32(c, 16)) + "";
                                        w.Append(asciiValue);
                                        pos = pos + 5;
                                    }
                                    else
                                    {
                                        // *** Invalid format
                                        dsErr = -51;
                                        dsErrMsg = "Expected \\u escape sequence to be followed by a valid four-digit hex value @Line:" + currLine(pos, ref js) + " @Position:" + currLinePos(pos, ref js) + " (e-51)";
                                        brk = true;
                                    }
                                    break;
                                default:
                                    // *** Invalid format
                                    dsErr = -52;
                                    dsErrMsg = "The \\ escape character is not followed by a valid value @Line:" + currLine(pos, ref js) + " @Position:" + currLinePos(pos, ref js) + " (e-52)";
                                    brk = true;
                                    break;
                            } //End switch
                            break;
                        default:
                            w.Append(c);
                            break;
                    } //End switch
                } // end if (c==endChar)
                pos = pos + 1;
            } while (!brk);
            return w.ToString();
        } // End

        private void DeserializeObject(bool keepSP, bool keepCM)
        {
            iesJSON j2;
            iesJSON jNew = null;
            int pStatus;
            StringBuilder getSpace = null;
            string Key = "";
            // *** Look for name:value pairs as a Dictionary object
            _jsonType = "object";

            System.Collections.Generic.List<Object> v = new System.Collections.Generic.List<Object>();
            endpos = endpos + 1; // *** IMPORTANT! Move past { symbol
            pStatus = 1;

            if (keepSP || keepCM) { getSpace = new StringBuilder(); }
            do
            {
                //j2=null;
                findnext(keepSP, keepCM, ref getSpace);

                // *** Check if we are past the end of the string
                if (endpos >= _jsonString.Length)
                {
                    StatusErr(-15, "Past end of string before we found the } symbol as the end of the object @Line:" + currLine(endpos) + " @Position:" + currLinePos(endpos) + " (e-15)");
                    _value = v;
                    if (keepSP || keepCM) { AddSpaceAndClear(_keep, "finalSpace", ref getSpace); } // Store the final space/comments somewhere
                    return;
                }

                switch (pStatus)
                {
                    case 1:
                        // *** Create new object now so we can add things like preSpace
                        jNew = new iesJSON();
                        jNew.UseFlexJson = UseFlexJson;
                        jNew.ALLOW_SINGLE_QUOTE_STRINGS = ALLOW_SINGLE_QUOTE_STRINGS;
                        if (keepSP || keepCM)
                        {
                            jNew.keepSpacingAndComments(keepSP, keepCM); // Setup _keep object
                            AddSpaceAndClear(jNew._keep, "preKey", ref getSpace);
                        }

                        // *** Get KEY: This MUST be a string with the name of the parameter
                        // NOTE: We throw this object away later!
                        // Space/comments before the key are preSpace
                        j2 = new iesJSON();
                        j2.UseFlexJson = UseFlexJson;
                        j2.ALLOW_SINGLE_QUOTE_STRINGS = ALLOW_SINGLE_QUOTE_STRINGS;
                        if (keepSP || keepCM)
                        {
                            j2.keepSpacingAndComments(keepSP, keepCM); // Setup 2 object
                        }

                        j2.Deserialize(_jsonString, endpos, true);

                        if (j2.Status != 0)
                        {
                            if (j2.Status == -11)
                            {
                                // *** This is actually legitimate.  All white space can indicate an empty object.  For example {}
                                // *** Here we are lenient and also allow a missing parameter.  For example {,
                                pStatus = 4;
                            }
                            else
                            {
                                StatusErr(-18, "Failed to find parameter name in parameter:value pair @Line:" + currLine(endpos) + " @Position:" + currLinePos(endpos) + " (e-18) [" + j2.StatusMessage + "]");
                            }
                        }
                        else
                        {
                            if (j2.jsonType != "string")
                            {
                                StatusErr(-19, "Parameter name in parameter:value pair must be a quoted string @Line:" + currLine(endpos) + " @Position:" + currLinePos(endpos) + " (-e19)");
                            }
                            else
                            {
                                Key = j2.ValueString;
                                endpos = j2.endpos;
                                pStatus = 2;
                            }
                        }
                        j2 = null;
                        break;
                    case 2:
                        // Expecting : symbol in value/pair combo
                        // Space/comments here are before the separator
                        if (keepSP || keepCM) { AddSpaceAndClear(jNew._keep, "postKey", ref getSpace); }
                        if (substr(_jsonString, endpos, 1) != ":")
                        {
                            StatusErr(-16, "Expected : symbol in parameter:value pair @Line:" + currLine(endpos) + " @Position:" + currLinePos(endpos) + " (e-16)");
                        }
                        else
                        {
                            endpos = endpos + 1;
                            pStatus = 3;
                        }
                        break;
                    case 3:
                        // Expecting value after the : symbol in value/pair combo
                        // Space/comments here are after the separator
                        // jNew should already be setup when we found the "key"
                        if (jNew == null)
                        {
                            StatusErr(-23, "System error. JSON object not initiated @Line:" + currLine(endpos) + " @Position:" + currLinePos(endpos) + " (e-23)");
                        }
                        else
                        {
                            // jNew should already be setup when we found the "key"
                            //jNew=new iesJSON();
                            //jNew.UseFlexJson=UseFlexJson;
                            //jNew.ALLOW_SINGLE_QUOTE_STRINGS=ALLOW_SINGLE_QUOTE_STRINGS;
                            if (keepSP || keepCM)
                            {
                                //jNew.keepSpacingAndComments(keepSP,keepCM); // Setup _keep object
                                AddSpaceAndClear(jNew._keep, "preSpace", ref getSpace);
                            }
                            jNew.Deserialize(_jsonString, endpos, true);
                            int newendpos = jNew.endpos;
                            // Check for blank=null (return status -11 and _jsonType=null)
                            if ((jNew.Status == -11) && (jNew.jsonType == "null") && (_UseFlexJson == true))
                            {
                                // Note: jNew.status=-11 indicates nothing was found where there should have been a value - for FLEX JSON this is legitimate.
                                //jNew=CreateItem(null);
                                jNew._status = 0; // this is OK
                            }
                            if (jNew.Status != 0)
                            {
                                StatusErr(-21, "Failed to find value in parameter:value pair. @Line:" + currLine(endpos) + " @Position:" + currLinePos(endpos) + " (e-21) [" + jNew.StatusMessage + "]");
                            }
                            else
                            {
                                // Note: Above, jNew.status=-11 indicates nothing was found where there should have been a value - for FLEX JSON this is legitimate.
                                // *** For all cases: object, array, string, number, boolean, or null
                                jNew.Parent = this;
                                jNew._key = Key;  // This must be done after serialization
                                v.Add(jNew);  // FUTURE: THIS IS WRONG!  WE MUST CHECK TO SEE IF THE KEY ALREADY EXISTS!
                                endpos = newendpos;
                                pStatus = 4;
                            }
                        }
                        //jNew=null; // This is needed later to store postSpace
                        break;
                    case 4:
                        // Expecting a , to indicate next element, or } to end object
                        // Space/comments here are after the value
                        if (keepSP || keepCM)
                        {
                            if (jNew != null) { AddSpaceAndClear(jNew._keep, "postSpace", ref getSpace); }
                            else { AddSpaceAndClear(_keep, "postSpace", ref getSpace); }
                        }
                        string c;
                        c = substr(_jsonString, endpos, 1);
                        if (c != "," && c != "}")
                        {
                            StatusErr(-17, "Expected , symbol to separate value pairs @Line:" + currLine(endpos) + " @Position:" + currLinePos(endpos) + " (e-17)");
                        }
                        else
                        {
                            endpos = endpos + 1;
                            // *** if we found } then successful completion of deserializing the object
                            if (c == "}") { _value = v; return; }
                            pStatus = 1;
                        }
                        jNew = null;  // Clear for next object
                        break;
                } //End Select
            } while (_status == 0); // Do Loop
            jNew = null;
            _value = v;
        } // End Deserialize_object()

        private void DeserializeArray(bool keepSP, bool keepCM)
        {
            iesJSON j2 = null;
            int pStatus;
            string c;
            StringBuilder getSpace = null;
            // *** Create an array of objects.  Each can be a value(string,number,boolean,object,array,null)
            _jsonType = "array";

            System.Collections.Generic.List<Object> v = new System.Collections.Generic.List<Object>();
            endpos = endpos + 1; // *** IMPORTANT! Move past [ symbol
            pStatus = 1;
            if (keepSP || keepCM) { getSpace = new StringBuilder(); }
            do
            {
                //j2 = null; // do not clear j2 between status 1 and 2
                findnext(keepSP, keepCM, ref getSpace);

                // *** Check if we are past the end of the string
                if (endpos > _jsonString.Length)
                {
                    StatusErr(-41, "Past end of string before we found the ] symbol as the end of the object @Line:" + currLine(endpos) + " @Position:" + currLinePos(endpos) + " (e-41)");
                    _value = v;
                    if (keepSP || keepCM) { AddSpaceAndClear(_keep, "finalSpace", ref getSpace); } // Store the final space/comments somewhere
                    return;
                }

                switch (pStatus)
                {
                    case 1:
                        // *** This should be a JSON item: object, array, string, boolean, number, or null
                        // Space/comments located here are preSpace
                        //if (keepSP || keepCM) { AddSpaceAndClear(_keep, "preSpace", ref getSpace); } // Wrong: This was prepending the space before the [ symbol
                        j2 = new iesJSON();
                        j2.UseFlexJson = UseFlexJson;
                        j2.ALLOW_SINGLE_QUOTE_STRINGS = ALLOW_SINGLE_QUOTE_STRINGS;
                        if (keepSP || keepCM)
                        {
                            j2.keepSpacingAndComments(keepSP, keepCM); // Setup _keep object
                            AddSpaceAndClear(j2._keep, "preSpace", ref getSpace);
                        }
                        j2.Deserialize(_jsonString, endpos, true);
                        if (j2.Status != 0)
                        {
                            if (j2.Status == -11)
                            {
                                // *** This is actually legitimate.  All white space can indicate an empty array.  For example []
                                // *** Here we are lenient and also allow a missing item.
                                // *** For example [,  NOTE! In this case, the missing item does NOT count as an element of the array!
                                // *** if you want to skip an item legitimately, use NULL  For example [NULL,
                                pStatus = 2;
                            }
                            else
                            {
                                StatusErr(-42, "Failed to find item in JSON array @Line:" + currLine(endpos) + " @Position:" + currLinePos(endpos) + " (e-42) [" + j2.StatusMessage + "]");
                            }
                        }
                        else
                        {
                            // *** For all cases: object, array, string, number, boolean, null - keep the entire JSON object.
                            j2.Parent = this;
                            v.Add(j2);
                            endpos = j2.endpos;
                            pStatus = 2;
                        }
                        //j2=null; // this object is needed below to store postSpace
                        break;
                    case 2:
                        // Expecting a , to indicate next element, or ] to end array
                        // Space/comments here are after the value
                        if (keepSP || keepCM)
                        {
                            if (j2 != null) { AddSpaceAndClear(j2._keep, "postSpace", ref getSpace); }
                            else { AddSpaceAndClear(_keep, "postSpace", ref getSpace); }
                        }
                        c = substr(_jsonString, endpos, 1);
                        if (c != "," && c != "]")
                        {
                            StatusErr(-43, "Expected , symbol to separate value in JSON array @Line:" + currLine(endpos) + " @Position:" + currLinePos(endpos) + " (e-43)");
                        }
                        else
                        {
                            endpos = endpos + 1;
                            // *** if we found ] then successful completion of deserializing the object
                            if (c == "]") { _value = v; return; }
                            pStatus = 1;
                        }
                        j2 = null;
                        break;
                } // End Select
            } while (_status == 0); //Loop
            j2 = null;
            _value = v;
        } // End

        private void AddSpaceAndClear(iesJSON toObj, string toParam, ref StringBuilder getSpace)
        {
            string newS = getSpace.ToString();
            if (!String.IsNullOrEmpty(newS))
            {
                toObj.AddToObjBase(toParam, toObj.GetStr(toParam) + newS);
            }
            getSpace.Length = 0;
        }

        private void AddSpace(iesJSON toObj, string toParam, string newSpace)
        {
            if (!String.IsNullOrEmpty(newSpace))
            {
                toObj.AddToObjBase(toParam, toObj.GetStr(toParam) + newSpace);
            }
        }

        // Store the current line/position into the _keep object
        private void AddLinePosition(iesJSON toObj, int AbsolutePosition)
        {
            toObj.AddToObjBase("AbsolutePosition", AbsolutePosition + 1);
            toObj.AddToObjBase("LineNumber", currLine(AbsolutePosition));
            toObj.AddToObjBase("LinePosition", currLinePos(AbsolutePosition));
        }

        // During the parse process, calculate the current line number
        private int currLine(int AbsolutePosition)
        {
            return currLine(AbsolutePosition, ref _jsonString);
        }

        private static int currLine(int AbsolutePosition, ref string js)
        {
            string noCarriageReturns = iesJSONutilities.Left(js, AbsolutePosition).Replace("\n", "");  // careful not to count \r\n as two lines
            int lineNumber = AbsolutePosition - noCarriageReturns.Length;
            return (lineNumber + 1);
        }

        // During the parse process, calculate the current position (character col) on current line
        private int currLinePos(int AbsolutePosition)
        {
            return currLinePos(AbsolutePosition, ref _jsonString);
        }

        private static int currLinePos(int AbsolutePosition, ref string js)
        {
            int lastPos = js.LastIndexOf("\n", AbsolutePosition);
            if (lastPos == -1)
            {
                return AbsolutePosition;
            }
            return (AbsolutePosition - lastPos);
        }

        private void findnext(bool keepSP, bool keepCM, ref StringBuilder keepSpace)
        {
            string c; int m;
            m = 0;
            while (endpos < _jsonString.Length)
            {
                c = substr(_jsonString, endpos, 1);
                if (c == "[")
                { // DEBUG 
                    string debug_c = substr(_jsonString, endpos, 100); // DEBUG
                    if (debug_c.IndexOf("Pulled") > 0)
                    {
                        c = c;
                    }
                } // DEBUG
                if (m == 0)
                {  // MODE 0 = outside of comment
                    if (c == "/")
                    {
                        if (substr(_jsonString, endpos + 1, 1) == "*")
                        {
                            m = 1;
                            endpos += 1;
                            if (keepCM) { keepSpace.Append("/*"); }
                        }
                        else if (substr(_jsonString, endpos + 1, 1) == "/")
                        {
                            m = 2;
                            endpos += 1;
                            if (keepCM) { keepSpace.Append("//"); }
                        }
                    }
                    if (m == 0)
                    {
                        if (!(c == " " || c == "\t" || c == "\n" || c == (Convert.ToChar(10) + "") || c == (Convert.ToChar(13) + ""))) { return; }
                        if (keepSP) { keepSpace.Append(c); }
                    }
                }
                else
                {
                    // MODE 1 OR 2 = inside comment
                    //if (m==1 && c=="*") {
                    if ((m == 1 || m == 2) && c == "*")
                    { // DEBUG TEMP REPLACE LINE ABOVE
                        if (substr(_jsonString, endpos + 1, 1) == "/")
                        {
                            m = 0;
                            endpos += 1;
                            if (keepCM) { keepSpace.Append("*/"); }
                        }
                    }
                    else
                    {
                        string capp = c;
                        if (m == 2 && (c == (Convert.ToChar(10) + "") || c == (Convert.ToChar(13) + "")))
                        {
                            m = 0;
                            // Check if we need to handle \r\n as a line feed.
                            string c2 = substr(_jsonString, endpos + 1, 1);
                            if (c == (Convert.ToChar(10) + "") && c2 == (Convert.ToChar(13) + "")) { capp += c2; endpos += 1; }
                        }
                        if (keepCM) { keepSpace.Append(capp); }
                        //if (keepCM) { keepSpace.Append("[" + capp + "]"); } // DEBUG TEMP REPLACE LINE ABOVE
                    }
                }
                endpos = endpos + 1;
            } //Loop
        } // End

        private void findnextChr(string iStr, int ptr)
        {
            string c; int m;
            m = 0;
            while (ptr < iStr.Length)
            {
                c = substr(iStr, ptr, 1);
                if (m == 0)
                {  // MODE 0 = outside of comment
                    if (c == "/")
                    {
                        if (substr(iStr, ptr + 1, 1) == "*") { m = 1; }
                        if (substr(iStr, ptr + 1, 1) == "/") { m = 2; }
                    }
                    if (m == 0)
                    {
                        if (!(c == " " || c == "\t" || c == (Convert.ToChar(10) + "") || c == (Convert.ToChar(13) + ""))) { return; }
                    }
                }
                else
                {
                    // MODE 1 or 2 = inside comment
                    if (m == 1)
                    {
                        // mode = 1
                        if (c == "*") { if (substr(iStr, ptr + 1, 1) == "/") { m = 0; ptr += 1; } }
                    }
                    else
                    { // mode = 2
                        if (c == (Convert.ToChar(10) + "") || c == (Convert.ToChar(13) + ""))
                        {
                            m = 0;
                            string c2 = substr(iStr, ptr + 1, 1);
                            if (c == (Convert.ToChar(10) + "") && c2 == (Convert.ToChar(13) + "")) { ptr += 1; }
                        }
                    }
                }
                ptr = ptr + 1;
            } //Loop
        } // End findnextChr()

        private void findnumber()
        {  // FUTURE: REMOVE THIS ROUTINE?  IS IT USED ANYMORE?
            string c; bool brk = false;
            do
            {
                if (endpos >= _jsonString.Length) { break; }
                c = substr(_jsonString, endpos, 1);
                switch (c)
                {
                    case "0":
                    case "1":
                    case "2":
                    case "3":
                    case "4":
                    case "5":
                    case "6":
                    case "7":
                    case "8":
                    case "9":
                    case "+":
                    case "-":
                    case "e":
                    case "E":
                    case ".":
                        break;
                    default:
                        return;
                }
                endpos = endpos + 1;
            } while (!brk);
        } // End

        private void findstring()
        {  // Ends with :,}] and stops at a /* comment */ if it is FLEX JSON
            string c; bool brk = false;
            do
            {
                if (endpos >= _jsonString.Length) { break; }
                c = substr(_jsonString, endpos, 1);
                if ((c == ":") || (c == ",") || (c == "}") || (c == "]") || c == "\n" || c == (Convert.ToChar(10) + "") || c == (Convert.ToChar(13) + "")) { return; }
                if (_UseFlexJson)
                {
                    if (c == "/")
                    {
                        if (substr(_jsonString, endpos + 1, 1) == "*") { return; }
                        if (substr(_jsonString, endpos + 1, 1) == "/") { return; }
                    }
                } //START of a comment
                endpos = endpos + 1;
            } while (!brk);
        } // End

        // FUTURE: REMOVE THIS ROUTINE?  IS IT USED ANYMORE?
        private void findcommentend()
        {  // Starts with /* and ends with */ if it is FLEX JSON (assume we already found the start)
            string c; bool brk = false;
            do
            {
                if (endpos >= _jsonString.Length) { break; }
                c = substr(_jsonString, endpos, 1);
                if (c == "*") { if (substr(_jsonString, endpos + 1, 1) == "/") { return; } } //END of a comment
                endpos = endpos + 1;
            } while (!brk);
        } // End

        static public bool IsHex4(string sVal)
        {
            int k;
            string c;
            if (sVal.Length != 4) { return (false); }
            for (k = 1; k < sVal.Length; k++)
            {
                c = substr(sVal, k, 1).ToUpper();
                switch (c)
                {
                    case "0":
                    case "1":
                    case "2":
                    case "3":
                    case "4":
                    case "5":
                    case "6":
                    case "7":
                    case "8":
                    case "9":
                    case "A":
                    case "B":
                    case "C":
                    case "D":
                    case "E":
                    case "F":
                        break;
                    default:
                        return (false);
                } //End switch
            } //Next
            return (true);
        } // End Function

        // SerializeMe() - Use this to Serialize the items that are already in the iesJSON object.
        // Return: 0=OK, -1=Error
        public int SerializeMe()
        {
            System.Text.StringBuilder s = new System.Text.StringBuilder();
            int i, k;
            string preKey = "", postKey = "";
            if (stats != null) { IncStats("stat_SerializeMe"); }
            if (_status != 0) { return -1; }
            if (!ValidateValue()) { return -1; }
            try
            {
                if (_keep != null)
                {
                    // Here we ignore keepSpacing/keepComments - these flags are only used during the deserialize process
                    s.Append(_keep.GetStr("preSpace")); // preSpace of overall object/array or item
                }
                switch (_jsonType)
                {
                    case "object":
                        s.Append("{");
                        i = 0;
                        foreach (object o in this)
                        {
                            iesJSON o2 = (iesJSON)o;
                            if (i > 0) { s.Append(","); }
                            k = o2.SerializeMe();
                            if (o2._keep != null)
                            {
                                // Here we ignore keepSpacing/keepComments - these flags are only used during the deserialize process
                                // preSpace and postSpace are already added during the SerializeMe() call above.  Here we add preKye and postKey.
                                preKey = o2._keep.GetStr("preKey");
                                postKey = o2._keep.GetStr("postKey");
                            }
                            if ((k == 0) && (o2.Status == 0)) { s.Append(preKey + "\"" + o2.Key.ToString() + "\"" + postKey + ":" + o2.jsonString); }
                            else
                            {
                                _status = -53;
                                AddStatusMessage("ERROR: Failed to serialize Object item " + i + " [err-53][" + o2.Status + ":" + o2.StatusMessage + "]");
                                return -1;
                            }
                            i++;
                        }
                        s.Append("}");
                        break;
                    case "array":
                        s.Append("[");
                        i = 0;
                        foreach (object o in this)
                        {
                            iesJSON o2 = (iesJSON)o;
                            if (i > 0) { s.Append(","); }
                            k = o2.SerializeMe();
                            if ((k == 0) && (o2.Status == 0)) { s.Append(o2.jsonString); }
                            else
                            {
                                _status = -52;
                                AddStatusMessage("ERROR: Failed to serialize Array item " + i + " [err-52]");
                                return -1;
                            }
                            i++;
                        }
                        s.Append("]");
                        break;
                    case "null":
                        s.Append("null");
                        break;
                    case "string":
                        s.Append("\"" + EncodeString(this.ValueString) + "\"");
                        break;
                    default:
                        s.Append(this.ValueString);
                        break;
                }
                if (_keep != null)
                {
                    s.Append(_keep.GetStr("postSpace") +
                        _keep.GetStr("finalSpace"));
                }
                _jsonString = s.ToString();
                _jsonString_valid = true;
            }
            catch
            {
                _jsonString = "";
                InvalidateJsonString(1);
                _status = -94;
                AddStatusMessage("ERROR: SerializeMe() failed to serialize the iesJSON object. [err-94]");
                return -1;
            }
            return 0;
        } // End

        // Used to copy an object/class/array without building the JSON string.
        // For example, capturing the Response.QueryString collection or Response.Form collection
        public void CopyValues(object objValue)
        {
            if (stats != null) { IncStats("stat_CopyValues"); }
            // *** Goes through the "serialize" process, but only fills in the _value and not the _jsonString
            this.Serialize(objValue, true, false);
        } // End

        // Use Newtonsoft.Json to convert Objects to JSON
        //public static string SerializeCustomObject(Object unserializedObject)
        //{
        //    return JsonConvert.SerializeObject(unserializedObject);
        //}

        // Use Newtonsoft.Json to convert Objects to JSON
        //public static T DeserializeCustomObject<T>(String serializedString)
        //{
        //    return JsonConvert.DeserializeObject<T>(serializedString);
        //}

        // *** Serialize() - Use this to Serialize an object/array/item that is NOT currently in the iesJSON object (such as Response.Querystring)
        // *** BuildValue=true causes this routine to not only generate a JSON string, but also fill in the
        // *** iesJSON.value property.  In most cases we want this so we can later manipulate the JSON object.
        // *** IF ALL YOU NEED IS THE JSON STRING, then you can set BuildValue=false to do less work and use less memory.
        // *** (Note: ByVal for objValue lets us work with the "pointer" to the object without disrupting the Parent's pointer to the same object)
        //DEFAULT-PARAMETERS
        //public void Serialize(object objValue) { Serialize(ref objValue, true, true); }
        //public void Serialize(object objValue, bool BuildValue) { Serialize(ref objValue, BuildValue, true); }
        //public void Serialize(object objValue, bool BuildValue, bool BuildString) {
        public void Serialize(object objValue, bool BuildValue = true, bool BuildString = true)
        {
            string t;
            if (stats != null) { IncStats("stat_Serialize"); }
            t = GetObjType(objValue);

            this.Clear(false, 1); // *** Clears _value and _jsonString (sets status=0) and also notifies Parent that we have changed
            if (substr(t, 0, 10) == "dictionary") { t = "dictionary"; }

            switch (t)
            {
                case "string":
                    this.SerializeString((string)objValue, BuildValue, BuildString);
                    break;
                case "integer":
                case "double":
                case "float":
                case "number":
                    this.SerializeNumber(objValue, BuildValue, BuildString);
                    break;
                case "boolean":
                case "bool":
                    this.SerializeBoolean((bool)objValue, BuildValue, BuildString);
                    break;
                case "null":
                    jsonString = "";
                    break;
                case "list":
                case "list(of object)":
                    this.SerializeList(objValue, BuildValue, BuildString);
                    break;
                case "dictionary":
                case "httpvaluecollection":
                    this.SerializeCollectionKeys(objValue, BuildValue, BuildString);
                    break;
                case "iesjson":
                    // If this is an iesJSON object, we still need to "copy" it.
                    this.SerializeVoidJSON((iesJSON)objValue, BuildValue, BuildString);
                    break;
                default:
                    StatusErr(-97, "Invalid parameter type - cannot serialize [" + t + "]");
                    break;
            } // end switch

        } // End

        private void SerializeVoidJSON(iesJSON oJSON, bool BuildValue, bool BuildString)
        {
            //Note: _value and _jsonString have already been cleared
            //Have JSON object serialize itself and give us the string...
            if (BuildString)
            {
                _jsonString = oJSON.jsonString;
                _jsonString_valid = true;  // *** This does not indicate valid syntax, only indicates that _jsonString is filled in.
            }
            if (BuildValue)
            {
                if ((oJSON.jsonType != "object") && (oJSON.jsonType != "array"))
                {
                    // Just need to copy this one node
                    _key = oJSON.Key;
                    _value = oJSON.Value;
                    _value_valid = true;
                    _jsonType = oJSON.jsonType;
                }
                else
                { // oJSON is  either an object or an array
                    System.Collections.Generic.List<Object> newList = new System.Collections.Generic.List<Object>();
                    foreach (object o in oJSON)
                    {
                        newList.Add(((iesJSON)o).Clone(BuildValue, BuildString));
                    }
                    _key = oJSON.Key;
                    _value = newList;
                    _value_valid = true;
                    _jsonType = oJSON.jsonType;
                }
            }
            if (oJSON.Status != 0)
            {
                // *** Something we did caused the oJSON object to become invalid
                this.Clear(false, 1);
                _status = -10;  // *** Indicate an invalid JSON result.
            }
        } // End SerializeVoidJSON

        public void SerializeString(string vString, bool BuildValue, bool BuildString)
        {
            //Note: _value and _jsonString have already been cleared
            if (BuildString)
            {
                _jsonString = "\"" + this.EncodeString(vString) + "\"";
                _jsonString_valid = true;
            }
            if (BuildValue)
            {
                _value = vString;
                _value_valid = true;
                _jsonType = "string";
            }
        } // End

        public void SerializeNumber(object vNum, bool BuildValue, bool BuildString)
        {
            //Note: _value and _jsonString have already been cleared
            if (BuildString)
            {
                _jsonString = vNum + "";
                _jsonString_valid = true;
            }
            if (BuildValue)
            {
                _value = vNum;
                _value_valid = true;
                _jsonType = "number";
            }
        } // End

        public void SerializeBoolean(bool vBool, bool BuildValue, bool BuildString)
        {
            //Note: _value and _jsonString have already been cleared
            string newVal = "";
            if (vBool)
            {
                newVal = "true";
            }
            else
            {
                newVal = "false";
            }
            if (BuildString)
            {
                _jsonString = newVal;
                _jsonString_valid = true;
            }
            if (BuildValue)
            {
                _value = newVal;
                _value_valid = true;
                _jsonType = "boolean";
            }
        } // End

        public void SerializeList(object vList, bool BuildValue, bool BuildString)
        {
            //Note: _value and _jsonString have already been cleared
            string s = "["; iesJSON j2; int cnt = 0;
            if (BuildValue)
            {
                _value = new System.Collections.Generic.List<Object>();
                _value_valid = true;
                _jsonType = "array";
            }
            foreach (object v in (System.Collections.Generic.List<Object>)vList)
            {
                j2 = new iesJSON();
                j2.UseFlexJson = UseFlexJson;
                j2.ALLOW_SINGLE_QUOTE_STRINGS = ALLOW_SINGLE_QUOTE_STRINGS;
                object v2 = v;
                j2.Serialize(v2, BuildValue, BuildString);
                if (j2.Status == 0)
                {
                    if (BuildString)
                    {
                        if (cnt > 0) { s = s + ","; }
                        s = s + j2.jsonString;
                    }
                    if (BuildValue)
                    {
                        j2.Parent = this;
                        ((System.Collections.Generic.List<Object>)_value).Add(j2);
                    }
                    cnt = cnt + 1;
                }
                else
                {
                    // *** raise error!
                    StatusErr(-94, "Failed to serialize object @ " + cnt + " (e-94) [" + j2.StatusMessage + "]");
                    break;
                }
            } //Next
            s = s + "]";
            if (BuildString)
            {
                _jsonString = s;
                _jsonString_valid = true;
            }
        } // End

        public void SerializeCollectionKeys(object vList, bool BuildValue, bool BuildString)
        {
            //Note: _value and _jsonString have already been cleared
            string s = "{", t; iesJSON j2; int cnt = 0;
            bool fProcessed = false;
            if (BuildValue)
            {
                _value = new System.Collections.Generic.List<Object>();
                _value_valid = true;
                _jsonType = "object";
            }
            if (vList == null) { return; } //Return with an empty object
            t = GetObjType(vList);

            // *************************************** DICTIONARY
            if (t.IndexOf("dictionary") >= 0)
            {
                fProcessed = true;
                System.Collections.Generic.Dictionary<string, object> vList2 = (System.Collections.Generic.Dictionary<string, object>)vList;
                //System.Collections.ObjectModel.Collection<object> vList2=(System.Collections.ObjectModel.Collection<object>) vList;
                foreach (System.Collections.Generic.KeyValuePair<string, object> kvp in vList2)
                {
                    string v = kvp.Key;
                    object o = kvp.Value;
                    j2 = new iesJSON();
                    j2.UseFlexJson = UseFlexJson;
                    j2.ALLOW_SINGLE_QUOTE_STRINGS = ALLOW_SINGLE_QUOTE_STRINGS;
                    j2.Serialize(o, BuildValue, BuildString);
                    j2.Key = v;
                    if (j2.Status == 0)
                    {
                        if (BuildString)
                        {
                            if (cnt > 0) { s = s + ","; }
                            s = s + "\"" + v + "\":" + j2.jsonString;
                        }
                        if (BuildValue)
                        {
                            j2.Parent = this;
                            ((System.Collections.Generic.List<Object>)_value).Add(j2);
                        }
                        cnt = cnt + 1;
                    }
                    else
                    {
                        // *** raise error!
                        StatusErr(-93, "Failed to serialize collection key @ " + cnt + " key=" + v + " (e-93) [" + j2.StatusMessage + "]");
                        break;
                    }
                } // Next
            }

            // *************************************** HTML FORM RESPONSE
            if (t.IndexOf("http") >= 0)
            {
                fProcessed = true;
                //System.Collections.Generic.Dictionary<string,object> vList2=(System.Collections.Generic.Dictionary<string,object>) vList;
                //System.Collections.ObjectModel.Collection<object> vList2=(System.Collections.ObjectModel.Collection<object>) vList;
                System.Collections.Specialized.NameValueCollection vList2 = (System.Collections.Specialized.NameValueCollection)vList;
                foreach (string v in vList2)
                {
                    object o = vList2[v];
                    j2 = new iesJSON();
                    j2.UseFlexJson = UseFlexJson;
                    j2.ALLOW_SINGLE_QUOTE_STRINGS = ALLOW_SINGLE_QUOTE_STRINGS;
                    j2.Serialize(o, BuildValue, BuildString);
                    j2.Key = v;
                    if (j2.Status == 0)
                    {
                        if (BuildString)
                        {
                            if (cnt > 0) { s = s + ","; }
                            s = s + "\"" + v + "\":" + j2.jsonString;
                        }
                        if (BuildValue)
                        {
                            j2.Parent = this;
                            ((System.Collections.Generic.List<Object>)_value).Add(j2);
                        }
                        cnt = cnt + 1;
                    }
                    else
                    {
                        // *** raise error!
                        StatusErr(-93, "Failed to serialize collection key @ " + cnt + " key=" + v + " (e-93) [" + j2.StatusMessage + "]");
                        break;
                    }
                } // Next
            }

            if (!fProcessed) { _status = -87; AddStatusMessage("ERROR: Failed to process collection.  Type not valid: '" + t + "' [err-87]"); }

            s = s + "}";
            if (BuildString)
            {
                _jsonString = s;
                _jsonString_valid = true;
            }
        } // End

        // *** splitstr()
        // *** Used to split a reference string such as "O\'Brien".23."color"
        // *** The above reference would get from a json object, the O'Brien parameter in an object, then lookup the 23rd record in the array,
        // *** and look up the "color" parameter in that object.  Although this split follows the conventions of a json string (for simplicity)
        // *** it should not start with a "[" symbol, and typically does not contain boolean, object, array, or null entries.
        //DEFAULT-PARAMETERS
        //public object splitstr(string strInput) { return splitstr(strInput, '.'); }
        //public object splitstr(string strInput, char Separator) {
        public System.Collections.Generic.List<Object> splitstr(string strInput, char Separator = '.')
        {
            iesJSON j2; int pStatus; int ptr; string c; System.Collections.Generic.List<object> s;
            // *** Create an array of objects.  Each can be a value(string,number,boolean,object,array,null)
            // *** but typically should only be 'string' or 'number' for the internal use within GetObj() and GetStr()
            s = new System.Collections.Generic.List<Object>();
            ptr = 0;
            pStatus = 1;
            while (ptr <= strInput.Length)
            {
                j2 = null;
                findnextChr(strInput, ptr);

                switch (pStatus)
                {
                    case 1:
                        // *** This should be a JSON item: object, array, string, boolean, number, or null
                        j2 = new iesJSON();
                        j2.UseFlexJson = UseFlexJson;
                        j2.ALLOW_SINGLE_QUOTE_STRINGS = ALLOW_SINGLE_QUOTE_STRINGS;
                        j2.Deserialize(strInput, ptr, true);
                        if (j2.Status == 0)
                        {
                            if ((j2.jsonType == "string") || (j2.jsonType == "number"))
                            {
                                // *** For all cases: Object, Array, String, Number, Boolean, or null
                                s.Add(j2.Value);
                                ptr = j2.endpos;
                                pStatus = 2;
                            }
                            else { pStatus = 3; } // ERROR
                        }
                        else { pStatus = 3; } //ERROR
                        break;
                    case 2:
                        c = substr(strInput, ptr, 1);
                        if (c == (Separator + ""))
                        {
                            ptr = ptr + 1;
                            pStatus = 1;
                        }
                        else { pStatus = 3; } //ERROR
                        break;
                } // switch
                if (pStatus == 3) { return null; } // *** Indicates an error in string format
            } //do
            j2 = null;
            return s;
        } // End Function

        // *** GetObj()
        // *** Parses a reference (string) to traverse the json object/array and returns the node or data element found.
        // *** returns null if not found (ie. if ANY node in the reference is not found while traversing the json tree)
        // *** Reference string can be in one of two formats... unquoted dot reference or quoted dot reference
        // *** Unquoted example: OBrien.23.color   (can not contain special characters in parameter names)
        // *** Quoted example: "O'Brien".23."color"  (MUST quote all parameter names in this format)
        // NOTE! If strReference="" then we return "this" iesJSON... which is useless unless we have a generic routine that needs to provide something if strReference is blank.
        // FUTURE: Faster method = boolean search if _sorted=true
        public iesJSON GetObj(string strReference)
        {
            object r = null; iesJSON v = null; iesJSON v2 = null; int safety;
            if (stats != null) { IncStats("stat_GetObj"); }
            // *** Make sure the _value is valid.
            if (!ValidateValue()) { return v; }

            if (strReference.Trim() == "") { return this; }

            v = this;
            try
            {
                if (strReference.IndexOf("\"") > -1)
                {
                    // *** quoted format
                    r = splitstr(strReference).ToArray();
                }
                else
                {
                    // *** unquoted format
                    r = strReference.Split('.');
                }
            }
            catch { }
            if (r == null) { return null; }
            safety = 999;
            foreach (string k in (string[])r)
            {
                if (v == null) { return null; }
                try
                {
                    v2 = null;
                    switch (v.jsonType)
                    {
                        case "array":
                            System.Collections.Generic.List<Object> z;
                            try
                            {
                                z = (System.Collections.Generic.List<Object>)v._value;
                                int i = int.Parse(k);
                                v2 = (iesJSON)z[i];
                            }
                            catch { v2 = null; }
                            break;
                        case "object":
                            System.Collections.Generic.List<Object> z2;
                            try
                            {
                                int j = v.IndexOfKey(k);
                                if (j >= 0)
                                {
                                    z2 = (System.Collections.Generic.List<Object>)v._value;
                                    v2 = (iesJSON)z2[j];
                                }
                                else { v2 = null; }
                            }
                            catch { v2 = null; }
                            break;
                        default:
                            // *** This is a problem... we are looking for a parameter (or array value), but the returned item is a String, Number, or Boolean
                            // *** return null
                            return null;
                    } //switch
                }
                catch { }
                if (v2 == null) { return null; }
                v = v2;
                safety = safety - 1;
                if (safety <= 0) { return null; }
            } //Next
            return v;
        } // End Function

        // *** GetParentObj()
        public iesJSON GetParentObj(string strReference, ref object finalKey)
        {
            object[] r = null; iesJSON v = null; iesJSON v2 = null; int cnt;
            if (stats != null) { IncStats("stat_GetParentObj"); }
            // *** Make sure the _value is valid.
            if (!ValidateValue()) { return v; }
            if (String.IsNullOrEmpty(strReference)) { return Parent; }  // return Parent object (or null if _parent is not set)

            v = this;
            try
            {
                if (strReference.IndexOf("\"") > -1)
                {
                    // *** quoted format
                    r = splitstr(strReference).ToArray();
                }
                else
                {
                    // *** unquoted format
                    r = strReference.Split('.');
                }
            }
            catch { }
            if (r == null) { return null; }
            if (r.Length <= 1)
            {
                if ((_jsonType != "object") && (_jsonType != "array")) { return null; } // reference is a KEY but we are not an Object/Array.
                finalKey = strReference;
                return this; // WE are the parent of a KEY value
            }
            cnt = 0;
            finalKey = "";
            foreach (string k in r)
            {
                if (v == null) { return null; }
                try
                {
                    v2 = null;
                    switch (v.jsonType)
                    {
                        case "array":
                            System.Collections.Generic.List<Object> z;
                            try
                            {
                                z = (System.Collections.Generic.List<Object>)v._value;
                                int i = int.Parse(k);
                                v2 = (iesJSON)z[i];
                            }
                            catch { v2 = null; }
                            break;
                        case "object":
                            System.Collections.Generic.List<Object> z2;
                            try
                            {
                                int j = v.IndexOfKey(k);
                                if (j >= 0)
                                {
                                    z2 = (System.Collections.Generic.List<Object>)v._value;
                                    v2 = (iesJSON)z2[j];
                                }
                                else { v2 = null; }
                            }
                            catch { v2 = null; }
                            break;
                        default:
                            // *** This is a problem... we are looking for a parameter (or array value), but the returned item is a String, Number, or Boolean
                            // *** return null
                            return null;
                    } //switch
                }
                catch { }
                v = v2;
                cnt++;
                if (cnt >= r.Length - 1)
                {
                    // Quit before trying to find the last part.
                    finalKey = r[r.Length - 1];
                    //tmpStatusMsg=tmpStatusMsg + "((" + finalKey + "~" + v._key + "~" + v.jsonString + "))";  DEBUG
                    return v;
                }
                if (cnt > 999) { return null; } // Safety
            } //Next
            return v;
        } // End Function

        // *** GetStr()
        // *** This works the same as GetObj(), but forces the return of a string.
        // *** if item returned is a number, it is converted to a string.
        // *** if item to be returned is an object, array, NULL, OR MISSING then the routine will return a blank string ""
        //DEFAULT-PARAMETERS
        //public string GetStr(string strReference) { return GetStr(strReference,""); }
        //public string GetStr(string strReference, string sDefault) {
        public string GetStr(string strReference, string sDefault = "")
        {
            iesJSON v = null; string ret;
            if (stats != null) { IncStats("stat_GetStr"); }
            ret = sDefault;
            v = GetObj(strReference);
            if (v != null) { ret = v.ValueString; }
            return ret;
        } // End Function

        // *** GetInt()
        // *** This works the same as GetObj(), but forces the return of an integer.
        // *** if the item is not found or is not a number, then nDefault is returned.  (WARNING - no error is thrown)
        //DEFAULT-PARAMETERS
        //public int GetInt(string strReference) { return GetInt(strReference, 0); }
        //public int GetInt(string strReference, int nDefault) {
        public int GetInt(string strReference, int nDefault = 0)
        {
            iesJSON v = null;
            if (stats != null) { IncStats("stat_GetInt"); }
            int gt = nDefault;
            v = GetObj(strReference);
            if (v == null) { return gt; }
            switch (v.jsonType)
            {
                case "string":
                    try { gt = int.Parse(v._value + ""); } catch { }
                    break;
                case "number":
                    try { gt = Convert.ToInt32((double)v._value); } catch { }
                    break;
                case "boolean":
                    try { gt = Convert.ToInt32((bool)v._value); } catch { }
                    break;
                    // default - leave gt as default
            }
            return gt;
        } // End Function

        // *** GetBool()
        // *** This works the same as GetObj(), but forces the return of a boolean.
        // *** if the item is not found or is not a boolean, then nDefault is returned.  (WARNING - no error is thrown)
        //DEFAULT-PARAMETERS
        //public int GetBool(string strReference) { return GetBool(strReference, 0); }
        //public int GetBool(string strReference, int nDefault) {
        public bool GetBool(string strReference, bool nDefault = false)
        {
            iesJSON v = null;
            if (stats != null) { IncStats("stat_GetBool"); }
            bool gt = nDefault;
            v = GetObj(strReference);
            if (v == null) { return gt; }
            switch (v.jsonType)
            {
                case "string":
                    try { gt = bool.Parse(v._value + ""); } catch { }
                    break;
                case "number":
                    try { gt = Convert.ToBoolean((double)v._value); } catch { }
                    break;
                case "boolean":
                    try { gt = (bool)v._value; } catch { }
                    break;
                    // default - leave gt as default
            }
            return gt;
        } // End Function

        // *** GetDouble()
        // *** This works the same as GetObj(), but forces the return of a double.
        // *** if the item is not found or is not a number, then nDefault is returned.  (WARNING - no error is thrown)
        //DEFAULT-PARAMETERS
        //public int GetDouble(string strReference) { return GetInt(strReference, 0); }
        //public int GetDouble(string strReference, double nDefault) {
        public double GetDouble(string strReference, double nDefault = 0)
        {
            iesJSON v = null;
            if (stats != null) { IncStats("stat_GetDouble"); }
            double gt = nDefault;
            v = GetObj(strReference);
            if (v == null) { return gt; }
            switch (v.jsonType)
            {
                case "string":
                case "number":
                case "boolean":
                    try { gt = double.Parse(v + ""); } catch { }
                    break;
                    // default - leave gt as default
            }
            return gt;
        } // End Function

        // *** GetJSONstr()
        // *** This works the same as GetObj(), but forces the return of a JSON string.
        // *** Normally this would be used to pull a JSON Array or Object out of the middle of a larger JSON Array or Object
        // *** in JSON String format.  GetJSONstr("Joe") is the equivalent of GetObj("Joe").jsonString, but handles cases
        // *** where the target is not found without throwing an error.
        //DEFAULT-PARAMETERS
        //public string GetJSONstr(string strReference) { return GetJSONstr(strReference,""); }
        //public string GetJSONstr(string strReference, string strDefault) {
        public string GetJSONstr(string strReference, string strDefault = "")
        {
            iesJSON v = null;
            if (stats != null) { IncStats("stat_GetJSONStr"); }
            string g = strDefault;
            v = GetObj(strReference);
            switch (v.jsonType)
            {
                case "string":
                case "number":
                case "boolean":
                    g = "\"" + v.ValueString + "\"";
                    break;
                case "iesJSON":
                    g = v.jsonString;
                    break;
            } //End switch
            return g;
        } // End Function

        // Remove an item from an ARRAY or OBJECT
        public bool RemoveAt(string strReference, int AtPosition)
        {
            iesJSON v = null;
            if (stats != null) { IncStats("stat_RemoveAt"); }
            v = GetObj(strReference);
            if (v == null) { return false; }
            return v.RemoveAtBase(AtPosition);
        }

        // Remove an item from the current object (assuming it is an ARRAY or OBJECT)
        public bool RemoveAtBase(int AtPosition)
        {
            if (stats != null) { IncStats("stat_RemoveAtBase"); }
            if (this._jsonType == "array" || this._jsonType == "object")
            {
                System.Collections.Generic.List<Object> a;
                try
                {
                    a = (System.Collections.Generic.List<Object>)this._value;
                    a.RemoveAt(AtPosition);
                    this.InvalidateJsonString(1);
                }
                catch { return false; }
            }
            else { return false; }
            return true;
        }

        // *** Add an Item to an array (at position - optional)
        //DEFAULT-PARAMETERS
        //public int AddToArray(string strReference, object oItem) { return AddToArray(strReference, oItem, -1); }
        //public int AddToArray(string strReference, object oItem, int atPosition) {
        public int AddToArray(string strReference, object oItem, int atPosition = -1)
        {
            iesJSON v = null; string t = "";
            if (stats != null) { IncStats("stat_AddToArray"); }
            int ret = -1; // *** Default=Error
            v = GetObj(strReference);
            if (v == null) { return ret; }
            //if (v.jsonType!="array") { return ret; }  // This check is done later in AddToArrayBase() - Can't add to Array if strReference is not an array!
            t = GetObjType(oItem);
            if (t == "iesjson")
            {
                // We were handed an iesJSON object, so just add it
                ret = v.AddToArrayBase((iesJSON)oItem, atPosition);
            }
            else
            {
                // We were handed something else, so put it into an iesJSON object and then add it
                iesJSON o = iesJSON.CreateItem(oItem, this.TrackStats);
                if (o != null) { ret = v.AddToArrayBase(o, atPosition); }
            }
            return ret;
        } // End Function

        // AddToArrayBase()
        // NOTE: Works for Arrays and Objects - Same functionality and InsertAt(iesJSON, index)
        // FUTURE: Danger here of adding the same Key twice for an Object. Or adding an object with no key.  Need to error check.
        //DEFAULT-PARAMETERS
        //public int AddToArrayBase(object oItem) { return AddToArrayBase(oItem, -1); }
        //public int AddToArrayBase(object oItem, int atPosition) {
        public int AddToArrayBase(iesJSON oJ, int atPosition = -1)
        {
            if (stats != null) { IncStats("stat_AddToArrayBase"); }
            int ret = -1;
            if (_status != 0) { return ret; }
            if ((_jsonType != "array") && (_jsonType != "object")) { return ret; } // *** Requires that this is a json "array" or json "object"
            oJ.Parent = this;
            oJ.InvalidateJsonString(1);  // Invalidates jsonString for ITEM AND for THIS
            try
            {
                if ((atPosition < 0) || (atPosition >= ((System.Collections.Generic.List<Object>)_value).Count))
                {
                    ((System.Collections.Generic.List<Object>)_value).Add(oJ);
                }
                else
                {
                    ((System.Collections.Generic.List<Object>)_value).Insert(atPosition, oJ);
                }
                ret = 0;
            }
            catch { }
            return ret;
        } // End Function

        // Generic Add of any type: string, int, etc. - Adds to BASE array/object
        public int AddItem(object oItem, int atPosition = -1)
        {
            string t = "";
            if (stats != null) { IncStats("stat_AddItem"); }
            t = GetObjType(oItem);
            if (t == "iesjson")
            {
                // We were handed an iesJSON object, so just add it
                return AddToArrayBase((iesJSON)oItem, atPosition);
            }
            else
            {
                // We were handed something else, so put it into an iesJSON object and then add it
                iesJSON o = iesJSON.CreateItem(oItem, this.TrackStats);
                if (o != null) { return AddToArrayBase(o, atPosition); }
            }
            return -1; // Error
        }

        // Add() - Generic Add() function
        // Value
        // If iesJSON instance is an OBJECT: Must specify two parameters... a NAME and VALUE.  json.Add("name","value");
        // If iesJSON instance is an ARRAY: This routine only uses the first parameter... a VALUE.  json.Add("value");
        // Value can be a string, int, bool, etc. - Adds to BASE array/object
        public int Add(object Param1, object Param2 = null)
        {
            iesJSON o;
            if (_jsonType == "object")
            {
                string t = GetObjType(Param1);
                if (t != "string") { return -12; }
                o = iesJSON.CreateItem(Param2, this.TrackStats);
                return AddToObjBase((string)Param1, o);
            }
            if (_jsonType == "array")
            {
                o = iesJSON.CreateItem(Param1, this.TrackStats);
                return AddToArrayBase(o);
            }
            return -1; // Error
        }

        // *** AddToObj()
        // *** strReference is a reference to a json "object" within the jSON tree (same syntax as specified for GetObj() above)
        // *** sParam is the name of a parameter in the json "object" that will be updated or added
        // *** fMode=3 Add or Update (which ever is needed)
        // ***       1 Add only   (error code if parameter already exists)
        // ***       2 Update only  (error code if parameter does not already exist)
        // ***       -1 REMOVE  (error code if parameter does not already exist)
        // *** NOTE! You can pass in a JSON object with the "Key" already set (make sParam="") OR include sParam which will override the Key
        // ***   For all other object types, sParam must be supplied and becomes the Key.
        //DEFAULT-PARAMETERS
        //public int AddToObj(string strReference,string sParam, object oItem) { return AddToObj(strReference, sParam, oItem, 3); }
        //public int AddToObj(string strReference,string sParam, object oItem, int fMode) {
        public int AddToObj(string strReference, string sParam, object oItem, int fMode = 3)
        {
            if (stats != null) { IncStats("stat_AddToObj"); }
            iesJSON v = null;
            int ret = -1; // *** Default=Error
            if (_status != 0) { return ret; }
            v = GetObj(strReference);
            if (v == null) { return ret; }
            //if (v.jsonType!="object") { return ret; }  // This check is done later in AddToObjBase() - Can't add to Object if strReference is not an object!
            ret = v.AddToObjBase(sParam, oItem, fMode);
            return ret;
        } // End Function

        // Same as above, but assumes that 'this' iesJSON is an 'object' and that we are adding the oItem to 'this'
        //DEFAULT-PARAMETERS
        //public int AddToObjBase(string sParam, object  oItem) { return AddToObjBase(sParam,oItem,3); }
        //public int AddToObjBase(string sParam, object  oItem, int fMode) {
        public int AddToObjBase(string sParam, object oItem, int fMode = 3)
        {
            if (stats != null) { IncStats("stat_AddToObjBase"); }
            string t = "";
            int ret = -1; // *** Default=Error
                          //if (this.jsonType!="object") { return ret; }  // This check is done later in AddToObjBase2() - Can't add to Object if strReference is not an object!
            t = GetObjType(oItem);

            if (t == "iesjson")
            {
                if (sParam != "") { ((iesJSON)oItem).Key = sParam; }
                ret = this.AddToObjBase2((iesJSON)oItem, fMode);
            }
            else
            {
                // We were handed something else, so put it into an iesJSON object and then add it
                iesJSON o = iesJSON.CreateItem(oItem, this.TrackStats);
                o.Key = sParam;
                if (o != null) { ret = this.AddToObjBase2(o, fMode); }
            }
            return ret;
        }


        public int AddToObjBase222(string sParam, object oItem, int fMode, ref string tmpStatusMsg2)
        {
            if (stats != null) { IncStats("stat_AddToObjBase"); }
            string t = "";
            int ret = -1; // *** Default=Error
                          //if (this.jsonType!="object") { return ret; }  // This check is done later in AddToObjBase2() - Can't add to Object if strReference is not an object!
            t = GetObjType(oItem);
            if (t == "iesjson")
            {
                if (sParam != "") { ((iesJSON)oItem).Key = sParam; }
                ret = this.AddToObjBase2((iesJSON)oItem, fMode);
            }
            else
            {
                // We were handed something else, so put it into an iesJSON object and then add it
                iesJSON o = iesJSON.CreateItem(oItem, this.TrackStats);
                o.Key = sParam;
                if (o != null) { ret = this.AddToObjBase333(o, fMode, ref tmpStatusMsg2); }
            }
            return ret;
        }


        private int AddToObjBase333(iesJSON oJ, int fMode, ref string tmpStatusMsg2)
        {
            if (stats != null) { IncStats("stat_AddToObjBase2"); }
            int ret = -1, k;
            System.Collections.Generic.List<Object> v;
            if (_status != 0) { return ret; }
            // *** Requires that this is a json "object"
            if (_jsonType != "object") { return ret; }
            if (oJ.Key == "") { return ret; }  //  Key must be set to a valid 'key' string
            k = this.IndexOfKey(oJ.Key);
            v = (System.Collections.Generic.List<Object>)this._value;
            if (k >= 0)
            {
                // *** Key already exists.
                if (fMode == 1) { return -11; } // *** Key already exists.
                try
                {
                    this.InvalidateJsonString(1);
                    if (fMode > 0)
                    {
                        v[k] = oJ; // Replace k-th element with new item
                    }
                    else
                    {
                        // ASSUME REMOVE ONLY
                        v.RemoveAt(k);
                    }
                    ret = 0;
                }
                catch { }
            }
            else
            {
                // *** Key does NOT exist
                if ((fMode == 2) || (fMode <= -1)) { return -12; } // *** Key does not exist.
                                                                   //try {
                this.InvalidateJsonString(1);
                v.Add(oJ);
                ret = 0;
                //} catch { }
            }
            return ret;
        } // End Function

        // AddToObjBase()
        // See above method AddToObj() for valid values of 'fMode'
        //DEFAULT-PARAMETERS
        //private int AddToObjBase2(object oJ) { return addToObjBase(sParam, oJ, 3); }
        //private int AddToObjBase2(object oJ, int fMode) {
        private int AddToObjBase2(iesJSON oJ, int fMode = 3)
        {
            if (stats != null) { IncStats("stat_AddToObjBase2"); }
            int ret = -1, k;
            System.Collections.Generic.List<Object> v;
            if (_status != 0) { return ret; }
            // *** Requires that this is a json "object"
            if (_jsonType != "object") { return ret; }
            if (oJ.Key == "") { return ret; }  //  Key must be set to a valid 'key' string
            k = this.IndexOfKey(oJ.Key);
            v = (System.Collections.Generic.List<Object>)this._value;
            oJ.Parent = this;
            oJ.InvalidateJsonString(1); // This invalidates the jsonString for the new item AND for THIS whether the update below works or not.
            if (k >= 0)
            {
                // *** Key already exists.
                if (fMode == 1) { return -11; } // *** Key already exists.
                try
                {
                    //this.InvalidateJsonString(1);  // accomplished above
                    if (fMode > 0)
                    {
                        v[k] = oJ; // Replace k-th element with new item
                    }
                    else
                    {
                        // ASSUME REMOVE ONLY
                        v.RemoveAt(k);
                    }
                    ret = 0;
                }
                catch { }
            }
            else
            {
                // *** Key does NOT exist
                if ((fMode == 2) || (fMode <= -1)) { return -12; } // *** Key does not exist.
                try
                {
                    //this.InvalidateJsonString(1);  // accomplished above
                    v.Add(oJ);
                    ret = 0;
                }
                catch { }
            }
            return ret;
        } // End Function

        public int RemoveFromObj(string strReference, string sParam)
        {
            if (stats != null) { IncStats("stat_RemoveFromObj"); }
            // return AddToObj(strReference, sParam, "", -1); // Old method
            iesJSON v = null;
            v = GetObj(strReference);
            if (v != null) {
                return v.RemoveFromBase(sParam);
            }
            return -2; // Invalid strReference
        } // End Function

        public int RemoveFromBase(string sParam)
        {
            if (stats != null) { IncStats("stat_RemoveFromBase"); }
            // return AddToObjBase(sParam, "", -1); // Old method

            // *** Requires that this is a json "object"
            if (_jsonType != "object") { return -3; }
            if (string.IsNullOrWhiteSpace(sParam)) { return -4; }  //  Key must be set to a valid 'key' string
            int k = this.IndexOfKey(sParam);
            if (k >=0) { 
                if (this.RemoveAtBase(k) == true) {
                    return 0; // OK
                }
                return -9; // other error
            }
            return -1; // Key not found
        } // End Function

        public bool Exists(string strReference)
        {
            iesJSON j = null;
            j = GetObj(strReference);
            if (j != null) { return true; }
            return false;
        } // End Function

        public string EncodeString(string vString)
        {
            string s;
            // *** NOTE: This is a simple encode.  It needs to be expanded in the future!
            s = vString;
            s = s.Replace("\\", "\\\\"); // **** NOTE: THIS MUST BE FIRST (so we do not double-escape the items below!)
            s = s.Replace("\t", "\\t");
            s = s.Replace("\n", "\\n");
            s = s.Replace("\r", "\\r");
            s = s.Replace("\"", "\\\"");
            if (ENCODE_SINGLE_QUOTES) { s = s.Replace("'", "\\'"); }
            return s;
        } // End Function

        private void StatusErr(int nErr, string strErr)
        {
            _status = nErr;
            AddStatusMessage(strErr);
        } // End

        static public string GetObjType(object o)
        {
            string n = "";
            if (o == null) { return ("null"); }
            try
            {
                n = o.GetType().Name.ToLower().Trim();
            }
            catch { n = "error"; }
            return n;
        }

        static public string substr(string tStr, int nStart, int len)
        {
            string s = "";
            if ((len <= 0) || (nStart >= tStr.Length)) { return (""); }
            if (tStr.Length < (nStart + len)) { return (tStr.Substring(nStart)); }
            try { s = tStr.Substring(nStart, len); return (s); } catch { }
            return ("");
        }

        static public string substr(string tStr, int nStart)
        {
            string s = "";
            if (nStart >= tStr.Length) { return (""); }
            try { s = tStr.Substring(nStart); return (s); } catch { }
            return ("");
        }

        public void CreateStats()
        {
            if (NoStatsOrMsgs) { return; }
            if (stats == null) { TrackStats = true; }
            if ((_status != 0) || (!_value_valid)) { return; } // Cannot track stats for child objects if the object is invalid.
                                                               // Do the same for all child objects
            if ((_jsonType == "object") || (_jsonType == "array"))
            {
                foreach (object o in this)
                    ((iesJSON)o).CreateStats();
            }
        }

        // ClearStats() - Clear current stats
        // ContinueToTrackStats=true leaves them ready for new stats,  false stops the tracking process
        //DEFAULT-PARAMETERS
        //public void ClearStats() { ClearStats(true); }
        //public void ClearStats(bool ContinueToTrackStats) {
        public void ClearStats(bool ContinueToTrackStats = true)
        {
            stats = null;
            if ((ContinueToTrackStats) && (!NoStatsOrMsgs)) { TrackStats = true; } else { TrackStats = false; }
            if ((_status != 0) || (!_value_valid)) { return; } // Cannot iterate through an array/object if it is not valid.
            if ((this._jsonType == "object") || (this._jsonType == "array"))
            {
                foreach (object o in this)
                {
                    if (!NoStatsOrMsgs)
                    {
                        ((iesJSON)o).ClearStats(ContinueToTrackStats);
                    }
                    else
                    {
                        ((iesJSON)o).ClearStats(false);
                        ((iesJSON)o).NoStatsOrMsgs = true;
                    }
                }
            }
        } // End clearStats()


        public void IncStats(string stat)
        {
            int i;
            if (NoStatsOrMsgs) { return; }
            if (stats == null) { return; }
            i = stats.GetInt(stat, 0); // If not found, gets a 0
            stats.AddToObjBase222(stat, i + 1, 3, ref tmpStatusMsg);  //overwrite existing value
        }

        // AddStatusMessage()
        // Set _statusMsg to the new message/error string
        // Also add the message to a log in stats object, if stats are being tracked.
        public void AddStatusMessage(string msg)
        {
            _statusMsg = msg;
            if (NoStatsOrMsgs) { return; }
            if (stats == null) { return; }

            // Here we should make the stats error messages an array of error messages.
            // Check if StatusMessages exists in stats
            if (!stats.Contains("StatusMessages"))
            {
                iesJSON k = CreateEmptyArray(); // Create new empty JSON array
                stats.AddToObjBase("StatusMessages", k);
            }
            else { stats.AddToArray("StatusMessages", msg); }
        }
        //DEFAULT-PARAMETERS
        //public int DeserializeStream(System.IO.StreamReader inStream) { return DeserializeStream(inStream,false); }
        //public int DeserializeStream(System.IO.StreamReader inStream, bool OkToClip) {
        public int DeserializeStream(System.IO.StreamReader inStream, bool OkToClip = false)
        {
            string line; // line2;
            int cnt = 0;
            bool keepSP = false;
            System.Text.StringBuilder strjson = new System.Text.StringBuilder();
            if (_keep != null) { keepSP = this.keepSpacing; }
            try
            {
                // We need to add a line feed to the BEGINNING of each line after line 0 to facilitate KeepSP and // Comments
                // (if we add to the end of each line, then a file that is opened/saved multiple times could grow longer.
                while ((line = inStream.ReadLine()) != null)
                {
                    // Handling comments is not in findnext and not here
                    //line2=line.TrimStart();
                    //if (substr(line2,0,2)!="//") { strjson.Append(line); }
                    if (cnt > 0) { strjson.Append("\n"); }
                    strjson.Append(line);
                    cnt++;
                }
                Deserialize(strjson.ToString(), 0, OkToClip);
            }
            catch { _status = -33; return _status; }
            return 0; //success
        }

        //DEFAULT-PARAMETERS
        //public int DeserializeFile(string FilePath) { return DeserializeFile(FilePath,false); }
        //public int DeserializeFile(string FilePath, bool OkToClip) {
        public int DeserializeFile(string FilePath, bool OkToClip = false)
        {
            //if ( !System.IO.File.Exists(FilePath) ) { _status=-31; return _status; }
            try
            {
                using (System.IO.StreamReader sr = new System.IO.StreamReader(FilePath))
                {
                    return DeserializeStream(sr, OkToClip);
                }
            }
            catch { _status = -32; return _status; }
        }

        // DeserializeFlexFile() - Deserialize a file using FLEX format (also allows flags to set how to treat spacing/commments)
        //   flag values: -1 leave default value, 0 Set to FALSE, 1 Set to TRUE
        public int DeserializeFlexFile(string FilePath, bool OkToClip = false, int spacing_flag = -1, int comments_flag = -1)
        {
            UseFlexJson = true;
            if ((spacing_flag >= 0) || (comments_flag >= 0)) { keepSpacingAndComments(spacing_flag, comments_flag); }
            return DeserializeFile(FilePath, OkToClip);
        }

        // WriteToFile()
        // If object is setup as FlexJSON and with KeepSpacing/KeepComments, then file will be written accordingly
        // Return: 0=ok, -1=error
        public int WriteToFile(string FilePath)
        {
            try
            {
                string outString = this.jsonString;
                if (_status == 0)
                {
                    if (File.Exists(FilePath)) { File.Delete(FilePath); }
                    File.WriteAllText(FilePath, outString);
                }
            }
            catch { return -1; }
            return 0;
        }

        public void Sort()
        {
            // Cannot sort an error object
            if (_status != 0)
            {
                return;
            }
            // For single items or Arrays/Objects with only one item, do nothing (not an error)
            if ((_jsonType != "array") && (_jsonType != "object"))
            {
                return;
            }
            if (Length <= 1)
            {
                return;
            }

            System.Collections.Generic.List<object> toSort = (System.Collections.Generic.List<object>)_value;
            toSort.Sort();
            InvalidateJsonString();
        }

        // Sort(iesJsonComparer)
        // Example 1:  
        //     iesJsonComparer jCompare = new iesJsonComparer(1); //Sort by second column (zero based numbering)
        //     jsonObject1.Sort(jCompare);
        //
        // Example 2:
        //     iesJsonComparer jCompare = new iesJsonComparer(new iesJSON("[\"Age\",\"Name\"]");
        //     jsonObject1.Sort(jCompare);
        //
        // Example 3:  
        //     iesJsonComparer jCompare = new iesJsonComparer("_key"); //Sort object by the key name of each item in the object
        //     jsonObject1.Sort(jCompare);
        //	  
        public void Sort(iesJsonComparer useComparer)
        {
            // Cannot sort an error object
            if (_status != 0)
            {
                return;
            }
            // For single items or Arrays/Objects with only one item, do nothing (not an error)
            if ((_jsonType != "array") && (_jsonType != "object"))
            {
                return;
            }
            if (Length <= 1)
            {
                return;
            }

            System.Collections.Generic.List<object> toSort = (System.Collections.Generic.List<object>)_value;
            toSort.Sort(useComparer);
            InvalidateJsonString();
        }

        int IComparable.CompareTo(object compare2)
        {
            //this.Parent.tmpStatusMsg += "[COMPARETO:" + compare2.GetType().ToString().ToLower() + "]"; // DEBUG DEBUG DEBUG
            if (compare2.GetType().ToString().ToLower() != "iesjsonlib.iesjson")
            {
                return 0; // not comparable - FUTURE: Allow compare of iesJSON to number/string if it is a number/string?
            }
            return this.CompareTo((iesJSON)compare2);
        }

        public int CompareTo(iesJSON compare2)
        {
            // First we determine sort order by jsonType (eg. Numbers always get sorted in front of strings)
            // NOTE: Uses the jsonTypeEnum to determine sort order by jsonType
            int eType1, eType2;
            eType1 = this.jsonTypeEnum;
            eType2 = compare2.jsonTypeEnum;
            if (eType1 < eType2)
            {
                return -1;
            }
            if (eType2 < eType1)
            {
                return 1;
            }
            // Items that are of the same type get compared
            switch (eType1)
            {
                case jsonTypeEnum_null: // NULL
                    return 0; // Two null values are always equal
                case jsonTypeEnum_number: // Number
                    double dbl1, dbl2;
                    dbl1 = this.ToDbl();
                    dbl2 = compare2.ToDbl();
                    return dbl1.CompareTo(dbl2);
                case jsonTypeEnum_string: // String
                    string str1, str2;
                    str1 = this.ToStr();
                    str2 = compare2.ToStr();
                    //this.Parent.tmpStatusMsg += "[COMPARE:" + str1 + ":" + str2 + ":" + str1.CompareTo(str2).ToString() + "]";  // DEBUG
                    //return str1.CompareTo(str2); // Case Sensitive Compare
                    return String.Compare(str1, str2, comparisonType: StringComparison.OrdinalIgnoreCase); // Case Insensitive Compare
                case jsonTypeEnum_boolean: // Boolean
                    bool bool1, bool2;
                    bool1 = this.ToBool();
                    bool2 = compare2.ToBool();
                    return bool1.CompareTo(bool2);
                default:
                    // All other types do not get compared: FUTURE: compare Arrays/Objects based on first item in each?
                    return 0;
            }
        }

    } //END Class - iesJSON

    public class iesJsonComparer : System.Collections.Generic.IComparer<object>
    {
        int CompareCount = 0; // 0=non-array/non-object, 1=single field, 2+ = multiple fields
        int CompareType = 0; // 0=non-array/non-object, 1=array (or object) specifies column number(s), 2=object only specifies parameter names, 3=object only sort by _key
        System.Collections.Generic.List<int> ListIdx;  // If CompareType=1, this is a list of IDX values (column numbers)
        System.Collections.Generic.List<string> ListParams; // If CompareType=2, this is a list of parameter names

        public iesJsonComparer()
        {
            CompareCount = 0;
            CompareType = 0;
        }

        public iesJsonComparer(int ColumnIdx)
        {
            iesJsonSetComparer(ColumnIdx);
        }

        void iesJsonSetComparer(int ColumnIdx)
        {
            CompareCount = 1;
            CompareType = 1;
            ListIdx = new System.Collections.Generic.List<int>();
            ListIdx.Add(ColumnIdx);
        }

        public iesJsonComparer(string Parameter)
        {
            iesJsonSetComparer(Parameter);
        }

        void iesJsonSetComparer(string Parameter)
        {
            if (Parameter.ToLower() == "_key")
            {
                CompareCount = 0;
                CompareType = 3; // Sort by _key name (objects only)
            }
            else
            {
                // Sort by specified field/column name									  
                CompareCount = 1;
                CompareType = 2;
                ListParams = new System.Collections.Generic.List<string>();
                ListParams.Add(Parameter);
            }
        }

        public iesJsonComparer(System.Collections.Generic.List<int> Idxs)
        {
            CompareCount = Idxs.Count;
            CompareType = 1;
            ListIdx = new System.Collections.Generic.List<int>();
            foreach (int idx in Idxs)
            {
                ListIdx.Add(idx);
            }
        }

        public iesJsonComparer(System.Collections.Generic.List<string> Parameters)
        {
            CompareCount = Parameters.Count;
            CompareType = 2;
            ListParams = new System.Collections.Generic.List<string>();
            foreach (string param in Parameters)
            {
                ListParams.Add(param);
            }
        }

        public iesJsonComparer(iesJSON config)
        {
            string configType = config.jsonType;
            if (config.Status != 0)
            {
                configType = "error";
            }
            switch (configType)
            {
                case "number":
                    this.iesJsonSetComparer(config.ToInt());
                    break;
                case "string":
                    this.iesJsonSetComparer(config.ToStr());
                    break;
                case "array":
                case "object":
                    string jType = config[0].jsonType.ToLower();
                    if (jType != "number" && jType != "string")
                    {
                        jType = "";
                        // FUTURE: Raise an error?
                        CompareCount = 0;
                        CompareType = 0;
                    }
                    else
                    {
                        if (jType == "number")
                        {
                            CompareType = 1;
                            ListIdx = new System.Collections.Generic.List<int>();
                        }
                        if (jType == "string")
                        {
                            CompareType = 2;
                            ListParams = new System.Collections.Generic.List<string>();
                        }

                        int iCount = 0;
                        foreach (iesJSON j in config)
                        {
                            iCount++;
                            if (jType == "number")
                            {
                                ListIdx.Add(j.ToInt());
                            }
                            if (jType == "string")
                            {
                                ListParams.Add(j.ToStr());
                            }
                        } // end foreach
                        CompareCount = iCount;

                    } // end else
                    break;

                default:
                    //FUTURE: Raise an error?
                    CompareCount = 0;
                    CompareType = 0;
                    break;
            }
        }

        int System.Collections.Generic.IComparer<object>.Compare(object x, object y)
        {
            string objType1, objType2;
            objType1 = x.GetType().ToString().ToLower();
            objType2 = y.GetType().ToString().ToLower();

            // We are only designed to compare iesJSON classes
            if (objType1 != "iesjsonlib.iesjson")
            {
                // FUTURE: Raise error?
                return 0;
            }
            if (objType2 != "iesjsonlib.iesjson")
            {
                // FUTURE: Raise error?
                return 0;
            }

            iesJSON xJson = (iesJSON)x;
            iesJSON yJson = (iesJSON)y;

            // Compare by _key first so that logic below can check for CompareCount==0
            if (CompareType == 3)
            {
                // _key compare - both items should be Object Items with a 'key' value
                string key1 = xJson.Key.ToLower();
                string key2 = yJson.Key.ToLower();
                return key1.CompareTo(key2);
            }
            if (CompareType == 0 || CompareCount == 0)
            {
                // Normal compare - not comparing based on a column or parameter
                return xJson.CompareTo(yJson);
            }

            if (CompareType == 1 && CompareCount > 0)
            {
                // Compare two arrays based on 1 or more column numbers
                foreach (int idx in ListIdx)
                {
                    iesJSON x2Json = xJson[idx];
                    iesJSON y2Json = yJson[idx];
                    int ret = x2Json.CompareTo(y2Json);
                    if (ret != 0)
                    {
                        return ret; // Return first column that is not equal.
                    }
                }
                return 0; // All columns were equal
            }

            if (CompareType == 2 && CompareCount > 0)
            {
                // Compare two arrays based on 1 or more column numbers
                foreach (string param in ListParams)
                {
                    iesJSON x2Json = xJson[param];
                    iesJSON y2Json = yJson[param];
                    int ret = x2Json.CompareTo(y2Json);
                    if (ret != 0)
                    {
                        return ret; // Return first column that is not equal.
                    }
                }
                return 0; // All columns were equal
            }
            return 0;
        }
    } // End Class iesJsonComparer

    public class iesCache
    {
        public iesJSON cache = null;
        public int MaxItems = 0;
        public double seqCounter = 1;

        // CONSTRUCTORS:
        public iesCache()
        {
            MaxItems = 100; // default max
        }
        public iesCache(int SetMaxItems)
        {
            MaxItems = SetMaxItems;
        }

        public void CreateCacheIfNeeded()
        {
            if (cache == null)
            {
                cache = new iesJSON("{}");
            }
        }

        public double NextSequence()
        {
            seqCounter += 1;
            return seqCounter;
            /*
                    DateTime centuryBegin = new DateTime(2001, 1, 1, 0, 0, 0, 0, System.DateTimeKind.Utc);
                    return (DateTime.Now.Ticks - centuryBegin.Ticks).ToString();
                    */
            /*
                    return (DateTime.UtcNow -
                    new DateTime(1970, 1, 1, 0, 0, 0, 0, System.DateTimeKind.Utc)).TotalSeconds;
            */
        }

        public iesJSON get(string token, bool updateTimestamp = true)
        {
            iesJSON ret = null;
            CreateCacheIfNeeded();
            ret = cache[token];
            if (ret != null)
            {
                if (ret.Status != 0) { ret = null; }
                else if (ret.jsonType == "null") { ret = null; }
            }
            if (ret == null)
            {
                ret = new iesJSON("{}");
                ret["_timestamp"].Value = this.NextSequence();
                cache[token] = ret;
            }
            else
            {
                if (updateTimestamp)
                {
                    ret["_timestamp"].Value = this.NextSequence();
                }
            }
            RemoveOldestIfNeeded();
            return ret;
        }

        public void store(string token, iesJSON item)
        {
            CreateCacheIfNeeded();
            item["_timestamp"].Value = this.NextSequence();
            cache[token] = item;
            RemoveOldestIfNeeded();
        }

        public void RemoveOldestIfNeeded()
        {
            if (cache == null) { return; }
            if (cache.Length > MaxItems) { RemoveOldest(); }
        }

        public void RemoveOldest()
        {
            try
            {
                if (cache == null) { return; }
                double oldestTimestamp = 0;
                string oldestToken = "";
                foreach (iesJSON i in cache)
                {
                    double d = i["_timestamp"].ToDbl(0);
                    if ((oldestTimestamp == 0) || (d < oldestTimestamp))
                    {
                        oldestTimestamp = d;
                        oldestToken = i.Key;
                    }
                }
                if (oldestToken != "") { cache.RemoveFromBase(oldestToken); }
            }
            catch { }
        }

        public void clear()
        {
            int Safety;
            Safety = MaxItems + 999;
            while (cache.Length > 0)
            {
                cache.RemoveAtBase(0);
                Safety = Safety - 1;
                if (Safety <= 0) { break; }
            }
        }

    } // END Class - iesCache

    // JSON Utilities
    public static class iesJSONutilities
    {

        // ************************************************************************************************************
        // **************** ReplaceTags()
        // **************** Replaces [[Tags]] with values from a iesJSON object.
        // **************** If a [[Tag]] is not found in tagValues then...
        // ****************   If SetNoMatchBlank=true Then the tag is replaced with ""
        // ****************   If SetNoMatchBlank=false Then the tag is left in the string.
        // ****************
        public static string ReplaceTags(string inputString, iesJSON tagValues, bool SetNoMatchBlank = true, string startStr = "[[", string endStr = "]]", int lvl = 0)
        {
            int charPosition = 0;
            int beginning = 0;
            int startPos = 0;
            int endPos = 0;
            StringBuilder data = new StringBuilder();

            // Safety - keep from causing an infinite loop.
            if (lvl++ > 99) { return inputString; }

            do
            {
                // Let's look for our tags to replace
                charPosition = inputString.IndexOf(startStr, endPos);
                if (charPosition >= 0)
                {
                    startPos = charPosition;
                    endPos = inputString.IndexOf(endStr, startPos);
                    if (endPos < 0)
                    {
                        // We did not find a matching end ]].  Break out of loop.
                        charPosition = -2;
                    }
                    else
                    {
                        // We found a match...
                        string tag = inputString.Substring(startPos + startStr.Length, (endPos - startPos) - endStr.Length);

                        string replacement = "";
                        // Check to see if tagValues contains a value for this field.
                        if (tagValues.Contains(tag))
                        {
                            // Yes it does.
                            replacement = tagValues[tag].ToStr();

                            // Check if our replacement string has [[tags]] that need to be replaced
                            int pt = replacement.IndexOf(startStr);
                            if (pt >= 0)
                            {
                                // Recursive call to replace [[tags]]
                                string replace2 = ReplaceTags(replacement, tagValues, SetNoMatchBlank, startStr, endStr, lvl);
                                replacement = replace2;
                            }
                        }
                        else
                        {
                            // No it does not contain tag
                            if (SetNoMatchBlank == false) { replacement = startStr + tag + endStr; }
                        }

                        data.Append(inputString.Substring(beginning, (startPos - beginning)));
                        data.Append(replacement);
                        beginning = endPos + endStr.Length;
                    }  // End if (endPos<0) else
                }  // End if (charPosition >= 0)
            } while (charPosition >= 0);

            if (beginning < (inputString.Length))
            {
                data.Append(inputString.Substring(beginning));
            }

            return data.ToString();
        } // END ReplaceTags()

        public static string Left(string val, int numChars)
        {
            if (val == null) return val;
            if (numChars <= 0) return "";
            if (numChars >= val.Length) return val;
            return val.Substring(0, numChars);
        }

        public static string GetParamStr(iesJSON parameterLibrary, string parameter, string defaultValue = "", bool tagReplace = true, bool SetNoMatchBlank = true, string startStr = "[[", string endStr = "]]")
        {
            string newValue = defaultValue;

            if (parameterLibrary.Contains(parameter))
            {
                newValue = parameterLibrary[parameter].ToStr(defaultValue);
            }

            if (tagReplace)
            {
                newValue = ReplaceTags(newValue, parameterLibrary, SetNoMatchBlank, startStr, endStr);
            }
            return newValue;
        }  // End GetParamStr()

    } // END class iesJSONutilities

} //END namespace

