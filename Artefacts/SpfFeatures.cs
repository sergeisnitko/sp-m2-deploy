using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SPMeta2.Definitions;
using SPMeta2.Enumerations;
using SPMeta2.Syntax.Default;

namespace SPF.Artefacts
{
    public static class SpfFeatures
    {
        public static FeatureDefinition DisableMinimalDownloadStrategy = BuiltInWebFeatures.MinimalDownloadStrategy.Inherit(def =>
        {
            def.Enable = false;
        });

    }
}
