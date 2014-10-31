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
        private OlSensitivity _Sensitivity = OlSensitivity.olNormal;

        private System.DateTime _Start = System.DateTime.Now;
        private System.DateTime _End = System.DateTime.Now;

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

    public class TrackingObj : TBObj
    {

        private DateTime _Start;
        private DateTime _End;
        private string _Subject;
        private List<string> _Categories;
        private AppointmentItem _RObj;

        public TrackingObj()
        {
            _Start = System.DateTime.Now;
            _End = System.DateTime.Now;
            this.Subject = "";
            this.Categories = new List<String>();
            this.RObj = null;
        }
        public TrackingObj(AppointmentItem AObj)
        {
            this.Subject = "";
            this.Categories = new List<String>();

            if (AObj != null)
            {
                _Start = AObj.Start;
                _End = AObj.End;
                _Subject = AObj.Subject;
                if (AObj.Subject != null)
                {
                    string[] _cats = AObj.Categories.Split(';');
                    foreach (string c in _cats)
                    {
                        _Categories.Add(c.Trim());
                    }
                }
                _RObj = AObj;
            }
        }
        public AppointmentItem RObj
        {
            get
            {
                return _RObj;
            }
            set
            {
                _RObj = value;
            }
        }
        public List<String> Categories
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
            set
            {
                _Subject = value;
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
        public void Update()
        {
            try
            {
                this.RObj.Subject = this.Subject;
                this.RObj.Save();
            }
            catch { }
        }
    }
}
