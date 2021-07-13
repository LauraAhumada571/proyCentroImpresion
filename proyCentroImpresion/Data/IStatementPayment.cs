using proyCentroImpresion.Models;
using System.Collections.Generic;


namespace proyCentroImpresion.Data
{
    interface IStatementPayment
    {
        List<Statementpayment> GetStatementPayment();
    }
}
