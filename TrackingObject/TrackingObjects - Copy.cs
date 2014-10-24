using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.Office.Interop.Outlook;
using System.Collections;
using System.Xml.Serialization;
using System.ComponentModel;

namespace TrackingObject
{
    [DataObjectAttribute]
    public class TrackingCol : TBObj, System.Collections.IEnumerable
    {
        private const OlDefaultFolders olFolderCalendar = OlDefaultFolders.olFolderCalendar;
        private List<TrackingObj> _TrackingCol;
        OlSensitivity _Sensitivity = OlSensitivity.olNormal;
        private Dictionary<string, ValuesObj> _Categories = new Dictionary<string, ValuesObj>();

        System.DateTime _Start = System.DateTime.Now;
        System.DateTime _End = System.DateTime.Now;

        int _Count = 0;

        public TrackingCol()
        {
            _TrackingCol = new List<TrackingObj>();
            _TrackingCol.Clear();
            _Start = _End.AddMonths(-1);

        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return new TrackingEnumerator(_TrackingCol);
        }

        [DataObjectMethod(DataObjectMethodType.Select)]
        public bool Extract()
        {
            bool _Done = true;
            _TrackingCol.Clear();
            try
            {
                Microsoft.Office.Interop.Outlook.Application OL = new Microsoft.Office.Interop.Outlook.Application();
                NameSpace oNS = OL.GetNamespace("MAPI");
                MAPIFolder oCalendar = oNS.Session.GetDefaultFolder(olFolderCalendar);
                Items oApps = oCalendar.Items;
                oApps.IncludeRecurrences = true;

                oApps.Sort("[Start]", true);

                oApps = oApps.Restrict("[Start] >= '" + _Start.ToShortDateString() + "' And [End] <= '" + _End.ToShortDateString() + "'");

                foreach (AppointmentItem Ap in oApps)
                {
                    if (Ap != null)
                    {
                        if ((Ap.Sensitivity == this.Sensitivity) && (Ap.UserProperties != null) && (Ap.Start >= _Start) && (Ap.End <= _End))
                        {
                            TrackingObj lto = new TrackingObj(Ap);
                            _TrackingCol.Add(lto);
                            _Count = _TrackingCol.Count;

                            foreach (KeyValuePair<string, string> kpv in lto.Categories)
                            {
                                ValuesObj cv;
                                try
                                {
                                    cv = _Categories[kpv.Key];
                                }
                                catch (KeyNotFoundException)
                                {
                                     cv = new ValuesObj(kpv.Value);
                                }
                                cv.Add(kpv.Value, 10);
                                _Categories[kpv.Key] = cv;
                            }
                        }
                    }
                }
            }
            catch { _Done = false; _TrackingCol.Clear(); }

            return (_Done);
        }

        public void Add(TrackingObj T)
        {
            try
            {
                _TrackingCol.Add(T);
            }
            catch { }
        }

        public void Remove(TrackingObj T)
        {
            try
            {
                _TrackingCol.Remove(T);
            }
            catch { }
        }

        public void Clear()
        {
            _TrackingCol.Clear();
        }

        public Dictionary<string, ValuesObj> Categories
        {
            get
            {
                return _Categories;
            }
            set
            {
                _Categories = value;
            }
        }

        public OlSensitivity Sensitivity
        {
            get
            {
                return _Sensitivity;
            }
            set
            {
                _Sensitivity = value;
            }
        }

        public int Count
        {
            get
            {
                return _Count;
            }
        }

        public string Start
        {
            get
            {
                return _Start.ToString("dd/MM/yyyy");
            }
            set
            {
                if (value.Length == 10)
                {
                    string[] va = value.Split('/');
                    if (va.Length == 3)
                    {
                        _Start = new DateTime(Int32.Parse(va[2]),Int32.Parse(va[1]),Int32.Parse(va[0]));
                    }
                }
                else
                {
                    _Start = new DateTime();
                }
            }
        }

        public string End
        {
            get
            {
                return _End.ToString("dd/MM/yyyy");
            }
            set
            {
                if (value.Length == 10)
                {
                    string[] va = value.Split('/');
                    if (va.Length == 3)
                    {
                        _End = new DateTime(Int32.Parse(va[2]), Int32.Parse(va[1]), Int32.Parse(va[0]));
                    }
                }
                else
                {
                    _End = new DateTime();
                }
            }
        }
    }

    public class TrackingEnumerator : System.Collections.IEnumerator
    {

        public List<TrackingObj> _TrackingCol;
        private int iPosition = -1;

        public TrackingEnumerator(List<TrackingObj> T)
        {
            _TrackingCol = T;
        }

        public void Reset()
        {
            iPosition = -1;
        }

        public object Current
        {
            get
            {
                return _TrackingCol[iPosition];
            }
        }

        public bool MoveNext()
        {
            iPosition++;
            if (iPosition >= _TrackingCol.Count)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

    }
///
/// <summary>
/// Represent a set of Values for a Categorie
/// </summary>
/// 
    public class ValuesObj : TBObj
    {
        private Dictionary<string, int> _Cumul = new Dictionary<string, int>();
        private Dictionary<string, int> _Occurence = new Dictionary<string, int>();
        private List<string> _Label = new List<string>();

        public ValuesObj(string l)
        {
            _Label.Clear();
            _Label.Add(l);
            _Cumul.Clear();
            _Cumul[l] = 0;
            _Occurence.Clear();
            _Occurence[l] = 0;
        }

        public ValuesObj(string l, int v)
        {
            _Label.Clear();
            _Label.Add(l);
            _Cumul.Clear();
            _Cumul[l] = v;
            _Occurence.Clear();
            _Occurence[l] = 1;

        }

        public int Add(string l,int v)
        {
            if (_Label.LastIndexOf(l) == -1)
            {
                _Label.Add(l);
                _Cumul[l] = 0;
                _Occurence[l] = 0;
            }
            _Occurence[l]++;
            _Cumul[l] += v;
            return (_Cumul[l]);
        }

        public int Sub(string label, int v)
        {
            if (_Occurence[label] != 0)
            {
                _Occurence[label]--;
                _Cumul[label] = _Cumul[label] - v;
            }
            return (_Cumul[label]);
        }

/// <summary>
/// Properties
/// </summary>

        public List<string> Labels
        {
            get
            {
                return _Label;
            }
        }
        public Dictionary<string, int> Cumul
        {
            get
            {
                return _Cumul;
            }
        }
        public Dictionary<string, int> Occurence
        {
            get
            {
                return _Occurence;
            }
        }
    }

    public class TrackingObj : TBObj
    {

        private Dictionary<string, string> _BizProp = new Dictionary<string, string>();
        private DateTime _Start;
        private DateTime _End;
        private string _Subject;
        private string _Categories;

        public TrackingObj()
        {
            _BizProp.Clear();
            _Start = System.DateTime.Now;
            _End = System.DateTime.Now;
            _Subject = "";
            _Categories = "";
        }

        public TrackingObj(AppointmentItem AObj)
        {
            _BizProp.Clear();
            _Start = System.DateTime.Now;
            _End = System.DateTime.Now;
            _Subject = "";

            if (AObj != null)
            {
                if (AObj.UserProperties.Count != 0)
                {
                    foreach (UserProperty _up in AObj.UserProperties)
                    {
                        if ((AObj.UserProperties[_up.Name] != null) && (_up.Name.Substring(0,4).ToLower() == "mssa"))
                        {
                            _BizProp[_up.Name] = AObj.UserProperties[_up.Name].Value.ToString();
                        }
                    }
                }
                _Start = AObj.Start;
                _End = AObj.End;
                _Subject = AObj.Subject;
                _Categories = AObj.Categories;
                //foreach( char c in AObj.Categories)
                //{
                //    _Categories.Add(c);
                //}
            }
        }

        public string GetValue(string Key)
        {
            return _BizProp[Key];
        }

        public void SetValue(string Key, string Value)
        {
            _BizProp[Key] = Value;
        }

        public Dictionary<string, string> Categories
        {
            get
            {
                return _BizProp;
            }
            set
            {
                _BizProp = value;
            }
        }

        public string Start
        {
            get
            {
                return _Start.ToString("dd/MM/yyyy hh:mm");
            }
            set
            {
                if (value.Length == 10)
                {
                    string[] va = value.Split('/');
                    if (va.Length == 3)
                    {
                        _Start = new DateTime(Int32.Parse(va[2]), Int32.Parse(va[1]), Int32.Parse(va[0]));
                    }
                }
                else
                {
                    _Start = new DateTime();
                }
            }
        }

        public string Subject
        {
            get
            {
                return _Subject;
            }
        }

        public string End
        {
            get
            {
                return _End.ToString("dd/MM/yyyy hh:mm");
            }
            set
            {
                if (value.Length == 10)
                {
                    string[] va = value.Split('/');
                    if (va.Length == 3)
                    {
                        _End = new DateTime(Int32.Parse(va[2]), Int32.Parse(va[1]), Int32.Parse(va[0]));
                    }
                }
                else
                {
                    _End = new DateTime();
                }
            }
        }
    }
}
