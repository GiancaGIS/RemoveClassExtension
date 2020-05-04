namespace StudioAT.ArcGIS.ArcCatalog.AddIn.RemoveClassExtension
{
    using ESRI.ArcGIS.Carto;
    using ESRI.ArcGIS.Catalog;
    using ESRI.ArcGIS.CatalogUI;
    using ESRI.ArcGIS.esriSystem;
    using ESRI.ArcGIS.Framework;
    using ESRI.ArcGIS.Geodatabase;
    using StudioAT.ArcGIS.ArcCatalog.AddIn.RemoveClassExtension.Generic;
    using System;
    using System.Runtime.InteropServices;
    using System.Windows.Forms;

    public class RemoveClassExtension : ESRI.ArcGIS.Desktop.AddIns.Button
    {
        IGxApplication pGxApp = null;

        public RemoveClassExtension()
        {
            this.pGxApp = ArcCatalog.Application as IGxApplication;
        }

        protected override void OnClick()
        {
            try
            {
                IGxSelection gxSelection = this.pGxApp.Selection; 
                
                if (gxSelection.Count < 1)
                {
                    return;
                }

                // Inizializzo le variabili per la progress bar...
                ITrackCancel trkCancel = null;
                IProgressDialogFactory prgFact = new ProgressDialogFactoryClass();
                IStepProgressor stepProgressor = null;
                IProgressDialog2 progressDialog = null;

                int intNumberFcConverted = 0;
                int cont = gxSelection.Count;
                IEnumGxObject enumGxObject = gxSelection.SelectedObjects;

                stepProgressor = prgFact.Create(trkCancel, 0);
                progressDialog = stepProgressor as IProgressDialog2;
                progressDialog.Description = "Removing extensions...";
                progressDialog.Title = "Removing extensions...";
                progressDialog.Animation = esriProgressAnimationTypes.esriProgressSpiral;
                progressDialog.ShowDialog();

                stepProgressor.MinRange = 0;
                stepProgressor.MaxRange = cont;
                stepProgressor.StepValue = 1;
                stepProgressor.Show();

                IGxObject pGxObject = enumGxObject.Next();

                while (pGxObject != null)
                {
                    if (!(pGxObject is IGxDataset))
                    {
                        return;
                    }

                    IGxDataset pGxDataset = pGxObject as IGxDataset;
                    if (pGxDataset == null)
                    {
                        return;
                    }

                    //if (((pGxObject as IGxDataset).Type) != esriDatasetType.esriDTFeatureClass)
                    //{
                    //    return;
                    //}

                    this.engine(pGxDataset, pGxObject, ref intNumberFcConverted);
                    pGxObject = enumGxObject.Next();
                    stepProgressor.Step();
                }

                stepProgressor.Message = "End";
                stepProgressor.Hide();
                progressDialog.HideDialog();

                MessageBox.Show($@"Extension removed for {intNumberFcConverted} objects!", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message.ToString());
            }

        }

        private void engine(IGxDataset pGxDataset, IGxObject pGxObject, ref int i)
        {
            try
            {
                IDataset dataset = pGxDataset.Dataset;
            }
            catch (COMException COMex)
            {
                if (COMex.ErrorCode == -2147467259)
                {
                    if (MessageBox.Show("Not found component register so I don't see the UID: do I have to remove class extension? Are you sure?", "Remove class extension", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        try
                        {
                            IGxObject gxObject = pGxObject.Parent;
                            if (!(gxObject is IGxDatabase2))
                            {
                                gxObject = gxObject.Parent;
                            }

                            IFeatureWorkspaceSchemaEdit featureWorkspaceSchemaEdit = ((gxObject as IGxDatabase2).Workspace) as IFeatureWorkspaceSchemaEdit;
                            featureWorkspaceSchemaEdit.AlterClassExtensionCLSID(pGxDataset.DatasetName.Name, null, null);

                            IObjectClassDescription featureClassDescription = new FeatureClassDescriptionClass();
                            featureWorkspaceSchemaEdit.AlterInstanceCLSID(pGxDataset.DatasetName.Name, featureClassDescription.InstanceCLSID);

                            MessageBox.Show("Class extension removed: success! Restart ArcCatalog!", "Remove Class Extension", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Error " + ex.Message);
                            return;
                        }
                    }
                }
                else
                {
                    MessageBox.Show("Error " + COMex.ErrorCode.ToString() + ": " + COMex.Message);
                }

                return;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
                return;
            }


            if (!(pGxDataset.Dataset is IClass))
            {
                return;
            }

            IClass pClass = pGxDataset.Dataset as IClass;

            if (pClass.EXTCLSID == null)
            {
                //MessageBox.Show("No class extension!");
                return;
            }
            else
            {

                IObjectClassDescription ocDescription = new AnnotationFeatureClassDescriptionClass();
                if ((pClass.EXTCLSID.Value.ToString() == ocDescription.ClassExtensionCLSID.Value.ToString()) && (pClass.EXTCLSID.SubType == ocDescription.ClassExtensionCLSID.SubType))
                {
                    MessageBox.Show("Class extension well-know: I don't remove it (annotation)!", "Remove Class Extension", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                ocDescription = new DimensionClassDescriptionClass();
                if ((pClass.EXTCLSID.Value.ToString() == ocDescription.ClassExtensionCLSID.Value.ToString()) && (pClass.EXTCLSID.SubType == ocDescription.ClassExtensionCLSID.SubType))
                {
                    MessageBox.Show("Class extension well-know: I don't remove it (dimension)!", "Remove Class Extension", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                if (MessageBox.Show(string.Format("Class extension: {0}: do I have to remove class extension?", pClass.EXTCLSID.Value.ToString()), "Remove Class Extension", MessageBoxButtons.YesNo) 
                    == DialogResult.Yes)
                {
                    if (clsGenericStudioAT.RemoveClsExt(pClass))
                    {
                        i++;
                        //MessageBox.Show("Class extension removed: success!", "Remove Class Extension", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        MessageBox.Show("Class extension removed: error! I have problem with exclusive schema lock ", "Remove Class Extension", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        protected override void OnUpdate()
        {
            if (this.pGxApp.Selection.Count < 0)
            {
                this.Enabled = false;
                return;
            }
            else
            {
                this.Enabled = true;
                return;
            }
        }

        private bool FeatureClassWellKnown(UID uid)
        {
            IObjectClassDescription ocDescription = new AnnotationFeatureClassDescriptionClass();
            if ((ocDescription.ClassExtensionCLSID.Value.ToString() == uid.Value.ToString()) && (ocDescription.ClassExtensionCLSID.SubType == uid.SubType))
            {
                return true;
            }

            ocDescription = new DimensionClassDescriptionClass();
            if ((ocDescription.ClassExtensionCLSID.Value.ToString() == uid.Value.ToString()) && (ocDescription.ClassExtensionCLSID.SubType == uid.SubType))
            {
                return true;
            }

            return false;
        }

    }

    
}
