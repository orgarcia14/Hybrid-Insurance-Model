function as = annuity(x,n,v,s,P,r)
    as = 0;
    for i=1:5
        asi = 0;
        for t=0:n-1
            tPx = survivalP(t,x,P);
            asi = asi + (v^t)*tPx(s,i);
        end
        as = as + asi*(1-r(i));
    end

