function BsI = IBenefit(x,n,v,s,P,I,m)
state = zeros(1,5);
state(s) = 1;
BsI = 0;
for k=0:n-1
    kPx = survivalP(k,x,P);
    BsI = BsI + (v^(m*(k+1)))*(state)*(kPx)*(I(:,x+k));
end

