import React from "react";

const SheetCard = ({ sheetNames }) => {
  if (!sheetNames) return null;
  console.log(sheetNames);
  return (
    <div>
      {sheetNames.map((item) => (
        <div key={item}>{item}</div>
      ))}
    </div>
  );
};

export default SheetCard;
