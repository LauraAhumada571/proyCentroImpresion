using proyCentroImpresion.Models;
using System.Collections.Generic;

namespace proyCentroImpresion.Data
{
    interface IDataExtractCollection
    {
        List<DataExtractCollection> GetDataExtractCollection(List<Extract> extract, List<Statementpayment> statementpayment, List<WithheldExtract> withheldextract, List<OutputType> outputtype);
    }
}
