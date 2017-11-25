CREATE TABLE SummaryData (
       FilePath    text not null,
       DocID       text not null,
       PatID       text not null,
       DocCode	   text not null,
       DocName	   text not null,
       DeptCode	   text not null,
       DeptName    text not null,
       CRUserID    text not null,
       CRUserName  text not null,
       CRDate      text not null,
       OPUserID    text not null,
       OPDate      text not null,
       PageNo      integer not null,
       TotalPaneNo integer not null,
       OutputFlg   text not null,
       NewFileName text not null
       );