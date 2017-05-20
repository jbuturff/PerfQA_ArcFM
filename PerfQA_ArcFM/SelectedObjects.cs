using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ESRI.ArcGIS.Geodatabase;

namespace PerfQA_ArcFM
{
    class SelectedObjects
    {
        private string sClassName;
        private int iClassID;
        private ISelectionSet pSelSet;

        public SelectedObjects(IFeatureClass pFC)
        {
            IDataset pDS;

            pDS = (IDataset)pFC;
            iClassID = pFC.FeatureClassID;
            sClassName = pDS.Name;
            pSelSet = pFC.Select(null, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionEmpty, pDS.Workspace);
        }

        public string ClassName
        {
	        get
	        {
	            return sClassName;
	        }
	        set
            {
	            sClassName = value;
	        }
        }

        public int ClassID
        {
            get
            {
                return iClassID;
            }
            set
            {
                iClassID = value;
            }
        }

        public ISelectionSet SelectionSet
        {
            get
            {
                return pSelSet;
            }
        }
    }
}

