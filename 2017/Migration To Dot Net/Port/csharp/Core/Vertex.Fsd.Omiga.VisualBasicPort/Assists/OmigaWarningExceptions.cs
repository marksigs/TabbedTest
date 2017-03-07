using System;
using System.Collections.Generic;
using System.Text;

namespace Vertex.Fsd.Omiga.VisualBasicPort.Assists
{
	/// <summary>
	/// Base class for Omiga warning exceptions. 
	/// </summary>
	/// <remarks>
	/// For each record in the MESSAGE table where MESSAGETYPE = 'Warning', if the warning needs to 
	/// be handled by a catch block, then derive a new class from <b>OmigaWarningException</b>.
	/// </remarks>
	public class OmigaWarningException : ErrAssistException
	{
		/// <summary>
		/// Creates a new <see cref="OmigaWarningException"/> object for a specified Omiga error number and parameters.
		/// </summary>
		/// <param name="errorNumber">The Omiga error number.</param>
		/// <param name="parameters">The parameters.</param>
		/// <remarks>
		/// The <paramref name="parameters"/> are used to replace any %s place holders in the 
		/// MESSAGETEXT column read from the database.
		/// </remarks>
		public OmigaWarningException(int errorNumber, params string[] parameters)
			: base(errorNumber, parameters)
		{
		}

		/// <summary>
		/// Indicates whether this exception object is an Omiga warning.
		/// </summary>
		/// <returns>Always returns <b>true</b>.</returns>
		public override bool IsWarning()
		{
			return true;
		}
	}

	// Missing XML comment for publicly visible type or member.
	#pragma warning disable 1591

	public sealed class PasswordExpiredException : OmigaWarningException
	{
		public PasswordExpiredException(params string[] parameters)
			: base(118, parameters)
		{
		}
	}

	public sealed class PurchasePriceException : OmigaWarningException
	{
		public PurchasePriceException(params string[] parameters)
			: base(167, parameters)
		{
		}
	}

	public sealed class ChangingAmountRequestedException : OmigaWarningException
	{
		public ChangingAmountRequestedException(params string[] parameters)
			: base(178, parameters)
		{
		}
	}

	public sealed class DeletingSelectedLoanComponentException : OmigaWarningException
	{
		public DeletingSelectedLoanComponentException(params string[] parameters)
			: base(187, parameters)
		{
		}
	}

	public sealed class ReservedMortgageProductException : OmigaWarningException
	{
		public ReservedMortgageProductException(params string[] parameters)
			: base(188, parameters)
		{
		}
	}

	public sealed class QuotationAffordabilityException : OmigaWarningException
	{
		public QuotationAffordabilityException(params string[] parameters)
			: base(204, parameters)
		{
		}
	}

	public sealed class ReinstatedQuotationException : OmigaWarningException
	{
		public ReinstatedQuotationException(params string[] parameters)
			: base(217, parameters)
		{
		}
	}

	public sealed class TotalCoverAmountException : OmigaWarningException
	{
		public TotalCoverAmountException(params string[] parameters)
			: base(230, parameters)
		{
		}
	}

	public sealed class UpdateApplicationCostsException : OmigaWarningException
	{
		public UpdateApplicationCostsException(params string[] parameters)
			: base(303, parameters)
		{
		}
	}

	public sealed class AllowableIncomeFactorException : OmigaWarningException
	{
		public AllowableIncomeFactorException(params string[] parameters)
			: base(307, parameters)
		{
		}
	}

	public sealed class BankWizardWarningException : OmigaWarningException
	{
		public BankWizardWarningException(params string[] parameters)
			: base(545, parameters)
		{
		}
	}

	public sealed class DependantException : OmigaWarningException
	{
		public DependantException(params string[] parameters)
			: base(553, parameters)
		{
		}
	}

	public sealed class OtherResidentException : OmigaWarningException
	{
		public OtherResidentException(params string[] parameters)
			: base(554, parameters)
		{
		}
	}

	public sealed class ODITransformerException : OmigaWarningException
	{
		public ODITransformerException(params string[] parameters)
			: base(4599, parameters)
		{
		}
	}

	public sealed class ODIConverterException : OmigaWarningException
	{
		public ODIConverterException(params string[] parameters)
			: base(4699, parameters)
		{
		}
	}

	public sealed class BatchPaymentStatusException : OmigaWarningException
	{
		public BatchPaymentStatusException(params string[] parameters)
			: base(4714, parameters)
		{
		}
	}

	public sealed class BatchPaymentMethodException : OmigaWarningException
	{
		public BatchPaymentMethodException(params string[] parameters)
			: base(4715, parameters)
		{
		}
	}

	public sealed class BatchPayeeBankException : OmigaWarningException
	{
		public BatchPayeeBankException(params string[] parameters)
			: base(4716, parameters)
		{
		}
	}

	public sealed class LTVChangedException : OmigaWarningException
	{
		public LTVChangedException(params string[] parameters)
			: base(7020, parameters)
		{
		}
	}

	public sealed class CalculatedOutstandingTermException : OmigaWarningException
	{
		public CalculatedOutstandingTermException(params string[] parameters)
			: base(7028, parameters)
		{
		}
	}

	public sealed class EsurvUnlockApplicationException : OmigaWarningException
	{
		public EsurvUnlockApplicationException(params string[] parameters)
			: base(8511, parameters)
		{
		}
	}

	public sealed class FirstTitleUnlockApplicationException : OmigaWarningException
	{
		public FirstTitleUnlockApplicationException(params string[] parameters)
			: base(8530, parameters)
		{
		}
	}

	public sealed class CriticalDataChangedException : OmigaWarningException
	{
		public CriticalDataChangedException(params string[] parameters)
			: base(8536, parameters)
		{
		}
	}

	public sealed class IssueOfferException : OmigaWarningException
	{
		public IssueOfferException(params string[] parameters)
			: base(8538, parameters)
		{
		}
	}

	public sealed class CreditCheckExpiredException : OmigaWarningException
	{
		public CreditCheckExpiredException(params string[] parameters)
			: base(8701, parameters)
		{
		}
	}

	public sealed class NoSubmissionRouteException : OmigaWarningException
	{
		public NoSubmissionRouteException(params string[] parameters)
			: base(8702, parameters)
		{
		}
	}
}
