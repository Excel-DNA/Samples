# ObjectHandler with refresh and selective updates

The idea of the sample is that the ObjectHandler.cs is generic code, and you customize it by making your own wrapper functions (like those in ExcelFunctions.cs) as well as some back-end data source behind the handles (like calls to a database or some large calculation, as implemented in DataService.cs). The ObjectHandler.cs has a HandleInfo class, which represents an object instance (uniquely defined by the ObjectType + Parameters) that can be modified or updated, and each update will create a new handle version (represented by the HandleIndex). Whenever a HandleInfo must be updated, the updating code calls HandleInfo.Update(...).

The HandleInfo object gets an IObserver hooked up when it is created - it will use the IObserver to tell Excel to update the cell (this pushes into the RTD Topic). For our purpose, we're only interested in calling the IObserver.OnNext() - this is what happens inside HandleInfo.Update(...).

When the IObserver is updated with the next value, it pushes to RTD, which eventually forces an update of the cells with those handles, and then all the dependencies.

In the sample there are two macros (with menu buttons) - one forces a refresh on all the handles, the other pretends to call the back-end (say a SQL Server), passing in the 'RowVersion' information and then updates those objects that have a new version coming back from the server. (In the sample I just pick every second object to 'update'.)
