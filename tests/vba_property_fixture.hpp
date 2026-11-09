#pragma once
#include <string>
inline const std::string nativePropertyFixture=R"VBA(Private backingValue As Long
Private reads As Long
Private indexedValues(1 To 3) As Long
Private retainedObject As Object
Private secretValue As Long
Public Property Get Stored() As Long
reads=reads+1
Stored=backingValue
End Property
Public Property Let Stored(ByRef value As Long)
backingValue=value
value=value+1
End Property
Public Property Get ReadCount() As Long
ReadCount=reads
End Property
Public Property Get Indexed(ByVal index As Long) As Long
Indexed=indexedValues(index)
End Property
Public Property Let Indexed(ByVal index As Long, value As Long)
indexedValues(index)=value
End Property
Public Property Get ObjectValue() As Object
Set ObjectValue=retainedObject
End Property
Public Property Set ObjectValue(value As Object)
Set retainedObject=value
End Property
Private Property Get Secret() As Long
Secret=secretValue
End Property
Private Property Let Secret(value As Long)
secretValue=value
End Property
Public Property Get Failure() As Long
Err.Raise 11
End Property
Public Function LocalProbe() As Long
Stored=9
Indexed(2)=Stored+3
Secret=5
Dim bag As Object
Set bag=New Scripting.Dictionary
bag.Add "entry",7
Set ObjectValue=bag
Dim copy As Object
Set copy=ObjectValue
LocalProbe=Indexed(2)+Secret+copy("entry")
End Function
Public Function FormProbe() As Long
Me.Stored=9
Me.Indexed(2)=Me.Stored+3
Dim bag As Object
Set bag=New Scripting.Dictionary
bag.Add "entry",7
Set Me.ObjectValue=bag
Dim copy As Object
Set copy=Me.ObjectValue
copy("entry")=8
FormProbe=Me.Indexed(2)+Me.ObjectValue("entry")
End Function
)VBA";
