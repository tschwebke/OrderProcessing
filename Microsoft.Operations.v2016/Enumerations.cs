using System;

namespace Microsoft.Operations
{
    /// <summary>
    /// A command/instruction which denotes how a subscription or email action should be treated.
    /// Used for logging to the database.
    /// </summary>
    public enum BlacklistResult
    {
        /// <summary>
        /// The email address was not found to be blacklisted, OR there is nothing currently
        /// preventing this from being sent.
        /// </summary>
        Allow,

        /// <summary>
        /// The email should not be sent because the address is blacklisted (with no exceptions)
        /// </summary>
        Block,

        /// <summary>
        /// Indicates that the item is blocked because an error (not specified) was encountered.
        /// </summary>
        BlockError,

        /// <summary>
        /// Based on noise, block if does not meet the required noise level.
        /// </summary>
        BlockNoise,

        /// <summary>
        /// We don't have a result or the test has not yet been conducted.
        /// </summary>
        Unknown,
    }

    /// <summary>
    /// Shorthand reference for which environment is being referred to. Normally useful for
    /// translating directly into a string value.
    /// </summary>
    public enum DeploymentEnvironment
    {
        /// <summary>
        /// 'No Environment specified' or 'None'
        /// </summary>
        NON,

        /// <summary>
        /// Security, Integration, Testing
        /// </summary>
        SIT,

        /// <summary>
        /// Developer environment - specifically instances mounted on individual developer machines.
        /// </summary>
        DEV,

        /// <summary>
        /// User Acceptance Testing
        /// </summary>
        UAT,

        /// <summary>
        /// Production instance
        /// </summary>
        PROD,

        Unknown,
        Production,
        Staging
    }

    /// <summary>
    /// Identifies the various stages/states associated with processing of files.
    /// </summary>
    public enum FileProcessingStatus
    {
        /// <summary>
        /// File has been uploaded, but no further action has been taken.
        /// </summary>
        Uploaded,

        /// <summary>
        /// File has been validated, according to the base processing logic.
        /// </summary>
        Validated,

        /// <summary>
        /// The file has some errors which prevents it from being used.
        /// </summary>
        Problematic
    }

    /// <summary>
    /// Bitwise operators for identifying particular 'flags' about the outgoing message. IMPORTANT
    /// NOTE: The balance of the chosen items is pretty important. Be careful about extending this
    /// logic as it has the potential to become unbalanced.
    /// </summary>
    [FlagsAttribute]
    public enum WorkItemDynamics
    {
        /// <summary>
        /// Base, every work item which is referenced here, has typically just been modified.
        /// </summary>
        IsModified = 0,

        /// <summary>
        /// The work item has been newly created. That generally means all the 'BEFORE' values are blank.
        /// </summary>
        NewWorkItem = 1,

        RecipientChangedThisWorkItem = 2,
        RecipientCreatedThisWorkItem = 4,
        RecipientHasOwnershipBefore = 8,
        RecipientHasOwnershipAfter = 16,

        // Anything past this point, is not actively used, however it might be useful reference if required.

        OwnershipChanges = 32,
        OwnershipTransferred = 64,
        OwnershipCreated = 128,
        OwnershipDestroyed = 256
    }
}