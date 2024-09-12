function kPx= survivalP(k,x,P)
I = eye(5);
if k==0
    kPx = I;
else
    kPx = I;
    for i=0:k-1
        kPx = kPx*P{x+i};
    end
end