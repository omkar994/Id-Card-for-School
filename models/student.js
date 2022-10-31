const mongoose=require('mongoose');
const Schema=mongoose.Schema;

const StudentSchema=new Schema({
    class:String,
    division:String,
    stduName:String,
    bdate:Date,
    bloodGrp:String,
    contactNo:Number,
    address:String,
    rollNo:Number
});

module.exports=mongoose.model('student', StudentSchema);