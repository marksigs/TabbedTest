/*
--------------------------------------------------------------------------------------------
Workfile:			OmigaErrorExceptions.cs
Copyright:			Copyright © 2007 Marlborough Stirling
Description:		Specific exceptions for the OMIGAERROR enum. 
					Only add exceptions to this file if they are for an OMIGAERROR enum value.
--------------------------------------------------------------------------------------------
.Net Specific History:

Prog	Date		Description
AS		29/05/2007	First version.
--------------------------------------------------------------------------------------------
*/
using System;
using System.Collections.Generic;
using System.Text;

namespace Vertex.Fsd.Omiga.VisualBasicPort.Assists
{
	/// <summary>
	/// Base class for Omiga error exceptions. 
	/// </summary>
	/// <remarks>
	/// For each record in the MESSAGE table where MESSAGETYPE = 'Error', if the error needs to 
	/// be handled by a catch block, then derive a new class from <b>OmigaErrorException</b>.
	/// </remarks>
	public class OmigaErrorException : ErrAssistException
	{
		/// <summary>
		/// Creates a new <see cref="OmigaErrorException"/> object for a specified <see cref="OMIGAERROR"/> and parameters.
		/// </summary>
		/// <param name="omigaError">The <see cref="OMIGAERROR"/>.</param>
		/// <param name="parameters">The parameters.</param>
		/// <remarks>
		/// The <paramref name="parameters"/> are used to replace any %s place holders in the 
		/// MESSAGETEXT column read from the database.
		/// </remarks>
		public OmigaErrorException(OMIGAERROR omigaError, params string[] parameters)
			: base(omigaError, parameters)
		{
		}

		/// <summary>
		/// Creates a new <see cref="OmigaErrorException"/> object for a specified <see cref="OMIGAERROR"/> and parameters.
		/// </summary>
		/// <param name="omigaError">The <see cref="OMIGAERROR"/>.</param>
		/// <param name="transactionAction">Indicates the action to take for any COM+ transaction.</param>
		/// <param name="parameters">The parameters.</param>
		/// <remarks>
		/// The <paramref name="parameters"/> are used to replace any %s place holders in the 
		/// MESSAGETEXT column read from the database.
		/// <para>
		/// If <paramref name="transactionAction"/> is TransactionAction.SetOnErrorType and the 
		/// exception is an Omiga system error, then if there is a COM+ transaction it will be set to 
		/// aborted. 
		/// If <paramref name="transactionAction"/> is TransactionAction.SetOnErrorType and the 
		/// exception is an Omiga application error, then if there is a COM+ transaction it will be set 
		/// to completed. 
		/// </para>
		/// </remarks>
		public OmigaErrorException(OMIGAERROR omigaError, TransactionAction transactionAction, params string[] parameters)
			: base(omigaError, transactionAction, parameters)
		{
		}

		/// <summary>
		/// Creates a new <see cref="OmigaErrorException"/> object for a specified Omiga error number and parameters.
		/// </summary>
		/// <param name="errorNumber">The Omiga error number.</param>
		/// <param name="parameters">The parameters.</param>
		/// <remarks>
		/// The <paramref name="parameters"/> are used to replace any %s place holders in the 
		/// MESSAGETEXT column read from the database.
		/// </remarks>
		public OmigaErrorException(int errorNumber, params string[] parameters)
			: base(errorNumber, parameters)
		{
		}

		/// <summary>
		/// Indicates whether this exception object is an Omiga warning.
		/// </summary>
		/// <returns>Always returns <b>false</b>.</returns>
		public override bool IsWarning()
		{
			return false;
		}
	}

	// Missing XML comment for publicly visible type or member.
	#pragma warning disable 1591

	public sealed class UnspecifiedErrorException : OmigaErrorException
	{
		public UnspecifiedErrorException(params string[] parameters)
			: base(OMIGAERROR.UnspecifiedError, parameters)
		{
		}
	}

	public sealed class NoAcceptedQuoteException : OmigaErrorException
	{
		public NoAcceptedQuoteException(params string[] parameters)
			: base(OMIGAERROR.NoAcceptedQuote, parameters)
		{
		}
	}

	public sealed class RecordNotFoundException : OmigaErrorException
	{
		public RecordNotFoundException(params string[] parameters)
			: base(OMIGAERROR.RecordNotFound, parameters)
		{
		}

		public RecordNotFoundException(TransactionAction transactionAction, params string[] parameters)
			: base(OMIGAERROR.RecordNotFound, transactionAction, parameters)
		{
		}
	}

    public sealed class XMLParserErrorException : OmigaErrorException
	{
		public XMLParserErrorException(params string[] parameters)
			: base(OMIGAERROR.XMLParserError, parameters)
		{
		}
	}

    public sealed class InvalidKeyStringException : OmigaErrorException
	{
		public InvalidKeyStringException(params string[] parameters)
			: base(OMIGAERROR.InvalidKeyString, parameters)
		{
		}
	}

    public sealed class NoAfterImagePresentException : OmigaErrorException
	{
		public NoAfterImagePresentException(params string[] parameters)
			: base(OMIGAERROR.NoAfterImagePresent, parameters)
		{
		}
	}

    public sealed class CommandFailedException : OmigaErrorException
	{
		public CommandFailedException(params string[] parameters)
			: base(OMIGAERROR.CommandFailed, parameters)
		{
		}
	}

    public sealed class MissingPrimaryTagException : OmigaErrorException
	{
		public MissingPrimaryTagException(params string[] parameters)
			: base(OMIGAERROR.MissingPrimaryTag, parameters)
		{
		}
	}

    public sealed class InvalidParameterException : OmigaErrorException
	{
		public InvalidParameterException(params string[] parameters)
			: base(OMIGAERROR.InvalidParameter, parameters)
		{
		}
	}

    public sealed class DuplicateKeyException : OmigaErrorException
	{
		public DuplicateKeyException(params string[] parameters)
			: base(OMIGAERROR.DuplicateKey, parameters)
		{
		}
	}

    public sealed class NoDataForCreateException : OmigaErrorException
	{
		public NoDataForCreateException(params string[] parameters)
			: base(OMIGAERROR.NoDataForCreate, parameters)
		{
		}
	}

    public sealed class NoDataForUpdateException : OmigaErrorException
	{
		public NoDataForUpdateException(params string[] parameters)
			: base(OMIGAERROR.NoDataForUpdate, parameters)
		{
		}
	}

    public sealed class ArrayLimitExceededException : OmigaErrorException
	{
		public ArrayLimitExceededException(params string[] parameters)
			: base(OMIGAERROR.ArrayLimitExceeded, parameters)
		{
		}
	}

    public sealed class NoRowsAffectedException : OmigaErrorException
	{
		public NoRowsAffectedException(params string[] parameters)
			: base(OMIGAERROR.NoRowsAffected, parameters)
		{
		}
	}

    public sealed class InvalidNoOfRowsException : OmigaErrorException
	{
		public InvalidNoOfRowsException(params string[] parameters)
			: base(OMIGAERROR.InvalidNoOfRows, parameters)
		{
		}
	}

    public sealed class InternalErrorException : OmigaErrorException
	{
		public InternalErrorException(params string[] parameters)
			: base(OMIGAERROR.InternalError, parameters)
		{
		}
	}

    public sealed class MissingFieldDescException : OmigaErrorException
	{
		public MissingFieldDescException(params string[] parameters)
			: base(OMIGAERROR.MissingFieldDesc, parameters)
		{
		}
	}

    public sealed class MissingTableDescException : OmigaErrorException
	{
		public MissingTableDescException(params string[] parameters)
			: base(OMIGAERROR.MissingTableDesc, parameters)
		{
		}
	}

    public sealed class MissingTypeDescException : OmigaErrorException
	{
		public MissingTypeDescException(params string[] parameters)
			: base(OMIGAERROR.MissingTypeDesc, parameters)
		{
		}
	}

    public sealed class MissingElementValueException : OmigaErrorException
	{
		public MissingElementValueException(params string[] parameters)
			: base(OMIGAERROR.MissingElementValue, parameters)
		{
		}
	}

    public sealed class MissingElementException : OmigaErrorException
	{
		public MissingElementException (params string[] parameters)
			: base(OMIGAERROR.MissingElement, parameters)
		{
		}
	}

    public sealed class MissingKeyDescException : OmigaErrorException
	{
		public MissingKeyDescException(params string[] parameters)
			: base(OMIGAERROR.MissingKeyDesc, parameters)
		{
		}
	}

    public sealed class MissingUpdateSetException : OmigaErrorException
	{
		public MissingUpdateSetException(params string[] parameters)
			: base(OMIGAERROR.MissingUpdateSet, parameters)
		{
		}
	}

    public sealed class MissingTableNameException : OmigaErrorException
	{
		public MissingTableNameException(params string[] parameters)
			: base(OMIGAERROR.MissingTableName, parameters)
		{
		}
	}

    public sealed class MissingXMLTableNameException : OmigaErrorException
	{
		public MissingXMLTableNameException(params string[] parameters)
			: base(OMIGAERROR.MissingXMLTableName, parameters)
		{
		}
	}

    public sealed class MissingKeyException : OmigaErrorException
	{
		public MissingKeyException(params string[] parameters)
			: base(OMIGAERROR.MissingKey, parameters)
		{
		}
	}

    public sealed class NoRowsAffectedByDeleteAllException : OmigaErrorException
	{
		public NoRowsAffectedByDeleteAllException(params string[] parameters)
			: base(OMIGAERROR.NoRowsAffectedByDeleteAll, parameters)
		{
		}
	}

    public sealed class InValidKeyValueException : OmigaErrorException
	{
		public InValidKeyValueException(params string[] parameters)
			: base(OMIGAERROR.InValidKeyValue, parameters)
		{
		}
	}

    public sealed class InValidDataTypeValueException : OmigaErrorException
	{
		public InValidDataTypeValueException(params string[] parameters)
			: base(OMIGAERROR.InValidDataTypeValue, parameters)
		{
		}
	}

    public sealed class InValidKeyException : OmigaErrorException
	{
		public InValidKeyException(params string[] parameters)
			: base(OMIGAERROR.InValidKey, parameters)
		{
		}
	}

    public sealed class NoFieldsFoundException : OmigaErrorException
	{
		public NoFieldsFoundException(params string[] parameters)
			: base(OMIGAERROR.NoFieldsFound, parameters)
		{
		}
	}

    public sealed class NoFieldItemFoundException : OmigaErrorException
	{
		public NoFieldItemFoundException(params string[] parameters)
			: base(OMIGAERROR.NoFieldItemFound, parameters)
		{
		}
	}

    public sealed class NoFieldItemNameException : OmigaErrorException
	{
		public NoFieldItemNameException(params string[] parameters)
			: base(OMIGAERROR.NoFieldItemName, parameters)
		{
		}
	}

    public sealed class NoComboTagValueException : OmigaErrorException
	{
		public NoComboTagValueException(params string[] parameters)
			: base(OMIGAERROR.NoComboTagValue, parameters)
		{
		}
	}

    public sealed class InvalidDateTimeFormatException : OmigaErrorException
	{
		public InvalidDateTimeFormatException(params string[] parameters)
			: base(OMIGAERROR.InvalidDateTimeFormat, parameters)
		{
		}
	}

    public sealed class NoBeforeImagePresentException : OmigaErrorException
	{
		public NoBeforeImagePresentException(params string[] parameters)
			: base(OMIGAERROR.NoBeforeImagePresent, parameters)
		{
		}
	}

    public sealed class ChildRecordsFoundException : OmigaErrorException
	{
		public ChildRecordsFoundException(params string[] parameters)
			: base(OMIGAERROR.ChildRecordsFound, parameters)
		{
		}
	}

	public sealed class InvalidSearchStringException : OmigaErrorException
	{
		public InvalidSearchStringException(params string[] parameters)
			: base(OMIGAERROR.InvalidSearchString, parameters)
		{
		}
	}

    public sealed class MTSNotFoundException : OmigaErrorException
	{
		public MTSNotFoundException(params string[] parameters)
			: base(OMIGAERROR.MTSNotFound, parameters)
		{
		}
	}

    public sealed class UnableToConnectException : OmigaErrorException
	{
		public UnableToConnectException(params string[] parameters)
			: base(OMIGAERROR.UnableToConnect, parameters)
		{
		}
	}

    public sealed class MissingParameterException : OmigaErrorException
	{
		public MissingParameterException(params string[] parameters)
			: base(OMIGAERROR.MissingParameter, parameters)
		{
		}
	}

    public sealed class XMLMissingElementException : OmigaErrorException
	{
		public XMLMissingElementException(params string[] parameters)
			: base(OMIGAERROR.XMLMissingElement, parameters)
		{
		}
	}

    public sealed class XMLMissingElementTextException : OmigaErrorException
	{
		public XMLMissingElementTextException(params string[] parameters)
			: base(OMIGAERROR.XMLMissingElementText, parameters)
		{
		}
	}

    public sealed class XMLMissingAttributeException : OmigaErrorException
	{
		public XMLMissingAttributeException(params string[] parameters)
			: base(OMIGAERROR.XMLMissingAttribute, parameters)
		{
		}
	}

    public sealed class XMLInvalidAttributeValueException : OmigaErrorException
	{
		public XMLInvalidAttributeValueException(params string[] parameters)
			: base(OMIGAERROR.XMLInvalidAttributeValue, parameters)
		{
		}
	}

    public sealed class XMLMissingPrimaryKeyException : OmigaErrorException
	{
		public XMLMissingPrimaryKeyException(params string[] parameters)
			: base(OMIGAERROR.XMLMissingPrimaryKey, parameters)
		{
		}
	}

    public sealed class XMLInvalidRequestNodeException : OmigaErrorException
	{
		public XMLInvalidRequestNodeException(params string[] parameters)
			: base(OMIGAERROR.XMLInvalidRequestNode, parameters)
		{
		}
	}

    public sealed class XMLAttributeAlreadyExistsException : OmigaErrorException
	{
		public XMLAttributeAlreadyExistsException(params string[] parameters)
			: base(OMIGAERROR.XMLAttributeAlreadyExists, parameters)
		{
		}
	}

    public sealed class ObjectNotCreatableException : OmigaErrorException
	{
		public ObjectNotCreatableException(params string[] parameters)
			: base(OMIGAERROR.ObjectNotCreatable, parameters)
		{
		}
	}

    public sealed class SchemaNotLoadedException : OmigaErrorException
	{
		public SchemaNotLoadedException(params string[] parameters)
			: base(OMIGAERROR.SchemaNotLoaded, parameters)
		{
		}
	}

    public sealed class SchemaParseErrorException : OmigaErrorException
	{
		public SchemaParseErrorException(params string[] parameters)
			: base(OMIGAERROR.SchemaParseError, parameters)
		{
		}
	}

    public sealed class NotImplementedException : OmigaErrorException
	{
		public NotImplementedException(params string[] parameters)
			: base(OMIGAERROR.NotImplemented, parameters)
		{
		}
	}

    public sealed class InvalidMessageNoException : OmigaErrorException
	{
		public InvalidMessageNoException(params string[] parameters)
			: base(OMIGAERROR.InvalidMessageNo, parameters)
		{
		}
	}

    public sealed class InvalidMessageQueueTypeException : OmigaErrorException
	{
		public InvalidMessageQueueTypeException(params string[] parameters)
			: base(OMIGAERROR.InvalidMessageQueueType, parameters)
		{
		}
	}

    public sealed class InvalidPLOMUserLevelException : OmigaErrorException
	{
		public InvalidPLOMUserLevelException(params string[] parameters)
			: base(OMIGAERROR.InvalidPLOMUserLevel, parameters)
		{
		}
	}
}
