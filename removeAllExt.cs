using ESRI.ArcGIS.Carto;
using ESRI.ArcGIS.Catalog;
using ESRI.ArcGIS.CatalogUI;
using ESRI.ArcGIS.Geodatabase;
using StudioAT.ArcGIS.ArcCatalog.AddIn.RemoveClassExtension.Generic;
using System;
using System.Windows.Forms;


namespace StudioAT.ArcGIS.ArcCatalog.AddIn.RemoveClassExtension
{
    public class removeAllExt : ESRI.ArcGIS.Desktop.AddIns.Button
    {
        IGxApplication pGxApp = null;
        IGxContentsView contentsView = null;

        public removeAllExt()
        {
            this.pGxApp = ArcCatalog.Application as IGxApplication;
            this.contentsView = this.pGxApp as IGxContentsView;
        }

        protected override void OnClick()
        {
            try
            {
                if (this.pGxApp.Selection.Count != 1)
                {
                    return;
                }             

                IGxObject pGxObject = this.pGxApp.SelectedObject;


                if (!(pGxObject is IGxDataset))
                {
                    return;
                }

                IGxDataset pGxDataset = pGxObject as IGxDataset;

                if (pGxDataset == null)
                {
                    return;
                }

                if (((pGxObject as IGxDataset).Type) != esriDatasetType.esriDTFeatureDataset)
                {
                    MessageBox.Show("Select a Feature Dataset!", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }

                IFeatureDataset fDataset = pGxDataset.Dataset as IFeatureDataset;
                IWorkspace2 workspace = fDataset.Workspace as IWorkspace2;
                IFeatureWorkspace featureWorkspace = workspace as IFeatureWorkspace;
                
                
                DialogResult dialogResult =
                MessageBox.Show($@"Remove custom extension for all Feature Classes in {fDataset.Name}", "Attention", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (dialogResult == DialogResult.Yes)
                {
                    IEnumDataset enumDataset = fDataset.Subsets;
                    IDataset dataset = enumDataset.Next();

                    IObjectClassDescription ocDescriptionAnnotation = new AnnotationFeatureClassDescriptionClass();
                    IObjectClassDescription ocDescriptionDimension = new DimensionClassDescriptionClass();

                    while (dataset != null)
                    {
                        if (dataset is IClass)
                        {
                            IClass pClass = dataset as IClass;
                            if (pClass.EXTCLSID != null)
                            {
                                if ((pClass.EXTCLSID.Value.ToString().ToLower() != ocDescriptionAnnotation.ClassExtensionCLSID.Value.ToString().ToLower())
                                    && (pClass.EXTCLSID.SubType != ocDescriptionAnnotation.ClassExtensionCLSID.SubType)
                                    && (pClass.EXTCLSID.Value.ToString().ToLower() != ocDescriptionDimension.ClassExtensionCLSID.Value.ToString().ToLower())
                                    && (pClass.EXTCLSID.SubType != ocDescriptionDimension.ClassExtensionCLSID.SubType))
                                {
                                    clsGenericStudioAT.RemoveClsExt(pClass);
                                }
                            }
                        }

                        dataset = enumDataset.Next();
                    }

                    MessageBox.Show("All Extension removed!", "Ok", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

            }
            catch (Exception errore)
            {
                MessageBox.Show(errore.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                throw;
            }
        }

        protected override void OnUpdate()
        {
            this.Enabled = true;
        }
    }
}
