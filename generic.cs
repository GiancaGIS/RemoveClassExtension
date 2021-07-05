using ESRI.ArcGIS.Geodatabase;


namespace StudioAT.ArcGIS.ArcCatalog.AddIn.RemoveClassExtension.Generic
{
    internal static class clsGenericStudioAT
    {
        public static bool RemoveClsExt(IClass classExtension)
        {
            ISchemaLock schemaLock = (ISchemaLock)classExtension;
            bool result = false;
            try
            {
                // Attempt to get an exclusive schema lock.
                schemaLock.ChangeSchemaLock(esriSchemaLock.esriExclusiveSchemaLock);

                // Cast the object class to the IClassSchemaEdit2 interface.
                IClassSchemaEdit2 classSchemaEdit = (IClassSchemaEdit2)classExtension;

                // Clear the class extension.
                classSchemaEdit.AlterClassExtensionCLSID(null, null);

                IObjectClassDescription featureClassDescription = new FeatureClassDescriptionClass();
                classSchemaEdit.AlterInstanceCLSID(featureClassDescription.InstanceCLSID);

                result = true;
            }
            catch
            {
                result = false;
            }
            finally
            {
                schemaLock.ChangeSchemaLock(esriSchemaLock.esriSharedSchemaLock);
            }

            return result;
        }
    }
}
