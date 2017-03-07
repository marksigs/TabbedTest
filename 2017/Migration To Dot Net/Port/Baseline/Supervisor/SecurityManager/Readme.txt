This component provides all the functionality for add/removing and retrieving permissions to users (supervisor).

There is a management interface which compiles into a .tlb type-library. When this interface is cast onto the SecurityManager object, it provides all the functionality required to manage security.

Without it, the Security.dll simply allows applications to check for a given user's access to a named resource.

Note on Compiling
=================

If the interface (to the Security.dll) has to be altered you must perform the following steps (due to cylclic project referencing!): -

1. In the Security.vbp, remove the reference to the SecurityMgtInterface.tlb type-library.

2. Comment-out the 'Implements ISecurityManager' line in the SecurityManager class.

3. Comment-out all the methods which implement this interface.

4. Build the component (and break binary compatibility if required).

5. Copy the new compiled dll into the 'Binary Compatible' folder.

6. Now alter any methods in the SecurityMgtInterface.vbp project which require it. Note: If the interface was broken in step 4 above, the project reference in SecurityMgtInterface.vbp which points to the Security.dll will need re-linking.

7. Compile the SecurityMgtInterface.vbp project (and break binary compatibility if required).

8. Copy the new compiled dll into the 'Binary Compatible' folder.

9. Unregister the compiled SecurityMgtInterface.dll file as applications will only need to register the .tlb file as the .dll contains no implementation.

10. Re-open the Security.vbp project and re-link the project reference to this new .tlb file.

11. uncomment out the lines which were commented in step 2 and 3 above.

12. Build the new Security.dll (adjusting any methods to match any changes in the ISecurityManager interface).