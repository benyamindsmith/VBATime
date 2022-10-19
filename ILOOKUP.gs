// Not sure if this works. Googlescript Equivalent of ILOOKUP
function ILOOKUP(lookup_value,lookup_range,return_range) {
  var lookup_val;

  for(let i =0; i<lookup_range.legnth; i++){
    let lookup_ind = i;
    let lookup_val = lookup_range[i];

    if(lookup_val <= lookup_value){
      return(return_range[lookup_ind]);
    } 
  }
}
