option explicit
' The QSort class sorts arrays using the quicksort algorithm.
' This code is a derived work from the code found at:
'	http://4guysfromrolla.com/webtech/012799-2.shtml
'
' This version can sort arrays of values or objects, using a user-supplied
' comparison function.

public const CMP_LESS = -1
public const CMP_EQU = 0
public const CMP_GREATER = 1

public const ORDER_ASC = 1
public const ORDER_DESC = -1

function qsort_cmp(a,b)
	if a < b then 
		qsort_cmp = CMP_LESS
	elseif a > b then 
		qsort_cmp = CMP_GREATER
	else 
		qsort_cmp = CMP_EQU
	end if
end function

function array_swap(values, i,j)
	dim temp: temp = values(i)
	values(i) = values (j)
	values(j) = temp
end function

function array_swapO(values, i,j)
	dim temp : set temp = values(i)
	set values(i) = values (j)
	set values(j) = temp
end function

class QSort
	private f_order
	private f_cmp
	private f_swap
	private valueSort
	
	private sub class_initialize
		f_order = ORDER_ASC
	end sub

	' Compare a and b, taking the sort-order into account
	' If DESC, reverse the results of the comparison	
	private function comp(a,b)
		comp = f_cmp(a,b) * f_order
	end function
	
	private sub QSort(byref values, loBound,hiBound)
		dim pivot,loSwap,hiSwap
	
	  ' Two items to sort
		if hiBound - loBound = 1 then
			if comp(values(loBound), values(hiBound)) = CMP_GREATER then
				f_swap values,loBound,hiBound
			end if
			exit sub
		end If
	
	  ' Three or more items to sort
		dim pivotIndex : pivotIndex = int((loBound + hiBound) / 2)
		
		if valueSort then
			pivot = values(pivotIndex)
		else
			set pivot = values(pivotIndex)
		end if
		
		f_swap values, pivotIndex, loBound
		
		loSwap = loBound + 1
		hiSwap = hiBound
	  
		do
			' Find the right loSwap
			while (loSwap < hiSwap) and (comp(values(loSwap), pivot) <> CMP_GREATER)
				loSwap = loSwap + 1
			wend
			' Find the right hiSwap
			while (comp(values(hiSwap), pivot) = CMP_GREATER)
				hiSwap = hiSwap - 1
			wend
			' Swap values if loSwap is less then hiSwap
			if loSwap < hiSwap then
				f_swap values, loSwap, hiSwap
			End If
		loop while loSwap < hiSwap
	  
		if valueSort then
			values(loBound) = values(hiSwap)
			values(hiSwap) = pivot
		else
			set values(loBound) = values(hiSwap)
			set values(hiSwap) = pivot
		end if
	  
		' Recursively call function
		' 2 or more items in first section
		if loBound < (hiSwap - 1) then QSort values, loBound, hiSwap-1
		' 2 or more items in second section
		if hiSwap + 1 < hibound then QSort values, hiSwap+1, hiBound
	End Sub

	public property set Compare(func)
		set f_cmp = func
	end property
	
    public property let Order(sortOrder)
        f_order = sortOrder
    end property

    private sub DetermineSortType(byref values)
        if IsEmpty(f_cmp) then
            valueSort = true
            set f_cmp = GetRef("qsort_cmp")
            set f_swap = GetRef("array_swap")
        elseif IsObject(values(LBound(values))) then   
            ' User defined object sorting
            valueSort = false
            set f_swap = GetRef("array_swap0")
        else
            ' User defined value sorting
            valueSort = true
            set f_swap = GetRef("array_swap")
        end if
    end sub

	public function Sort(byref values)
        ' Don't sort empty arrays or arrays with only 1 value
        if UBound(values) < 1 then 
            Sort = values
            exit function
        end if
        
        DetermineSortType values
			
		QSort values, LBound(values), UBound(values)
    	Sort = values

		if valueSort then f_cmp = Empty
	end function
end class
